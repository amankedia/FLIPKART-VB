VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   10440
      Width           =   10575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   9840
      Width           =   5295
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
      Height          =   735
      Left            =   13680
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   12960
      Picture         =   "Formnew.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Search"
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
      Left            =   10080
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Search Any Product"
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
      Left            =   6720
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   6120
      Picture         =   "Formnew.frx":1487
      Stretch         =   -1  'True
      Top             =   480
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   3120
      Picture         =   "Formnew.frx":1FE2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Label1"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
