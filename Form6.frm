VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00C0FFFF&
   Caption         =   "PAYMENT "
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18705
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   18705
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11160
      TabIndex        =   2
      Top             =   8280
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000011&
      Caption         =   "CONFIRM"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   1
      Top             =   8280
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "OR,"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9120
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Form6.frx":2EAA9
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2895
      Left            =   10800
      TabIndex        =   3
      Top             =   2280
      Width           =   7695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "PLEASE SCAN THE QR CODE "
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Top             =   5760
      Width           =   7275
   End
   Begin VB.Image Image1 
      Height          =   4920
      Left            =   1440
      Picture         =   "Form6.frx":2EB32
      Stretch         =   -1  'True
      Top             =   840
      Width           =   7275
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub COMMAND1_CLICK()
MsgBox "YOUR TICKET IS BOOKED, THANK YOU "
Form3.Show
End Sub

Private Sub Command2_Click()
MsgBox "YOUR PAYMENT IS NOT COMPLETED,TICKET NOT BOOKED"
Form5.Show
Form6.Hide
End Sub

