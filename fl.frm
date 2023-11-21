VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0080FF80&
   Caption         =   "Form3"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18450
   LinkTopic       =   "Form3"
   Picture         =   "fl.frx":0000
   ScaleHeight     =   9765
   ScaleWidth      =   18450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "END"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      TabIndex        =   3
      Top             =   7800
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BOOK NEXT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   2
      Top             =   7800
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      Height          =   7575
      Left            =   1080
      Top             =   240
      Width           =   12735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Visit Again "
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   9720
      TabIndex        =   1
      Top             =   6840
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Thank You for using our service"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   7305
      Left            =   1200
      Picture         =   "fl.frx":1ED62
      Top             =   360
      Width           =   12465
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub Command2_Click()
End
End Sub
