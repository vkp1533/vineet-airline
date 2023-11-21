VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFF00&
   Caption         =   "one way booking"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Amita"
      Size            =   26.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   8550
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   7320
      TabIndex        =   7
      Top             =   6600
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   7320
      TabIndex        =   6
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   7320
      TabIndex        =   5
      Top             =   4080
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      TabIndex        =   4
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   3
      Top             =   8160
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   7215
      Left            =   960
      Top             =   2280
      Width           =   10815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER YOUR DETAILS :-"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1215
      Left            =   1200
      TabIndex        =   8
      Top             =   2520
      Width           =   6855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AGE :-"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   2520
      TabIndex        =   2
      Top             =   5400
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "HOST  NAME :-"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   2520
      TabIndex        =   1
      Top             =   4080
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ID NO. :-"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   6600
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command3_Click()
Form6.Show
Form2.Hide
End Sub

Private Sub Form_Load()
con.Open "provider= microsoft.jet.oledb.4.0;data source=F:\air ticker booking files\vi.mdb;Persist security info ="
rs.Open "select * from swayambhu", con, adOpenDynamic, adLockPessimistic
End Sub
Sub display()
Text2.Text = rs!Text2
Text3.Text = rs!Text3
Text4.Text = rs!Text4
End Sub
Private Sub Command2_Click()
Form2.Hide
o1.Show
End Sub
Private Sub COMMAND1_CLICK()
rs("name").Value = Text2.Text
rs("age").Value = Text3.Text
rs("idnum").Value = Text4.Text
MsgBox "Data Has Been Saved !!TICKET BOOKED", vbInformation, "Flight Booked "
rs.Update
Form3.Show
End Sub

