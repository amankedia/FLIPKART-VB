VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   1410
   ClientWidth     =   19080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0FFC0&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleMode       =   0  'User
   ScaleWidth      =   19080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer7 
      Interval        =   350
      Left            =   9000
      Top             =   3960
   End
   Begin VB.Timer Timer6 
      Interval        =   350
      Left            =   10320
      Top             =   4080
   End
   Begin VB.Timer Timer5 
      Interval        =   350
      Left            =   9000
      Top             =   3960
   End
   Begin VB.Timer Timer4 
      Interval        =   350
      Left            =   11760
      Top             =   6600
   End
   Begin VB.Timer Timer3 
      Interval        =   350
      Left            =   12480
      Top             =   6600
   End
   Begin VB.Timer Timer1 
      Interval        =   350
      Left            =   12120
      Top             =   6600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   6330
      Left            =   4440
      TabIndex        =   0
      Top             =   1440
      Width           =   9000
      Begin VB.Timer Timer2 
         Interval        =   350
         Left            =   6960
         Top             =   5160
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   735
         Left            =   2760
         TabIndex        =   2
         Top             =   6000
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1296
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Image Image5 
         Height          =   975
         Left            =   7920
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3960
         TabIndex        =   6
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Aman Kedia(CSE-155)"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   615
         Left            =   1920
         TabIndex        =   5
         Top             =   5040
         Width           =   4575
      End
      Begin VB.Image Image2 
         Height          =   1575
         Left            =   0
         Picture         =   "frmSplash.frx":04CC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3885
      End
      Begin VB.Image Image9 
         Height          =   1320
         Left            =   2400
         Picture         =   "frmSplash.frx":1D52
         Top             =   2640
         Width           =   1320
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Goudy Stout"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   4680
         TabIndex        =   4
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Goudy Stout"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   3480
         TabIndex        =   3
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   855
         Left            =   3960
         TabIndex        =   1
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   1320
         Index           =   1
         Left            =   2400
         Picture         =   "frmSplash.frx":29FC
         Top             =   2640
         Width           =   1320
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "A product of Aman Inc."
      BeginProperty Font 
         Name            =   "Kozuka Mincho Pr6N H"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   360
      Width           =   3735
   End
   Begin VB.Image Image4 
      Height          =   1455
      Left            =   14280
      Picture         =   "frmSplash.frx":3685
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   3525
   End
   Begin VB.Image Image3 
      Height          =   1695
      Left            =   840
      Picture         =   "frmSplash.frx":4F0B
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2805
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
Image9.Visible = Not Image9.Visible
End Sub

Private Sub Timer2_Timer()
Label3.Caption = Label3.Caption + 1
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = 100 Then
Timer2.Enabled = False
ProgressBar1.Value = 0
Unload Me
Flipkart.Show
End If


End Sub

Private Sub Timer3_Timer()
Label5.Visible = Not Label5.Visible
End Sub

Private Sub Timer4_Timer()
Image2.Left = Image2.Left + 50
End Sub

Private Sub Timer5_Timer()
Image3.Top = Image3.Top + 50

End Sub

Private Sub Timer6_Timer()
Image4.Top = Image4.Top - 50
End Sub

Private Sub Timer7_Timer()
Label6.Left = Label6.Left + 100
End Sub
