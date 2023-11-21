VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   10140
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18315
   BeginProperty Font 
      Name            =   "Amita"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   10140
   ScaleWidth      =   18315
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   6720
      Top             =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GO BACK"
      Height          =   975
      Left            =   8040
      TabIndex        =   3
      Top             =   4320
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   915
      Left            =   4440
      TabIndex        =   2
      Text            =   "SELECT PAYMENT MODE"
      Top             =   2640
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BOOK NOW"
      Height          =   975
      Left            =   4680
      TabIndex        =   1
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label3 
      Height          =   735
      Left            =   4560
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Round Trip Booking Details:-"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   1215
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   12975
   End
   Begin VB.Shape Shape1 
      Height          =   6615
      Left            =   4320
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub COMMAND1_CLICK()
If Combo1 = "QR SCAN" Then
Form6.Show
Else
p1f1.Show
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label3 = Date
Label2 = Time
Combo1.AddItem "QR SCAN"
Combo1.AddItem "UPI TRANSFER"
End Sub


Private Sub Label4_Click()

End Sub
