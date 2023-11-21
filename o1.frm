VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form o1 
   BackColor       =   &H00FF80FF&
   Caption         =   "One Way Trip "
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17025
   BeginProperty Font 
      Name            =   "Amita"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   Picture         =   "o1.frx":0000
   ScaleHeight     =   9750
   ScaleWidth      =   17025
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo2 
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   10560
      TabIndex        =   30
      Text            =   "select class"
      Top             =   6120
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<<<"
      Height          =   735
      Left            =   14040
      TabIndex        =   29
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">>>"
      Height          =   735
      Left            =   12000
      TabIndex        =   28
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add"
      Height          =   735
      Left            =   9840
      TabIndex        =   27
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "save"
      Height          =   615
      Left            =   6360
      TabIndex        =   26
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cancel"
      Height          =   615
      Left            =   3840
      TabIndex        =   25
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "confirm"
      Height          =   615
      Left            =   1440
      TabIndex        =   24
      Top             =   8400
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Height          =   735
      Left            =   15240
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   735
      Left            =   12360
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   735
      Left            =   12360
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   12360
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   5000
      TabIndex        =   17
      Top             =   6240
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   4995
      TabIndex        =   12
      Top             =   7080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      _Version        =   393216
      Format          =   162004993
      CurrentDate     =   45045
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   5000
      TabIndex        =   11
      Top             =   5400
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   5000
      TabIndex        =   10
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   5000
      TabIndex        =   9
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   5000
      TabIndex        =   8
      Top             =   2880
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Text            =   "select your flight"
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Left            =   2640
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Select Class Cabin "
      Height          =   1335
      Left            =   9000
      TabIndex        =   31
      Top             =   5640
      Width           =   6135
   End
   Begin VB.Shape Shape4 
      Height          =   1215
      Left            =   9240
      Top             =   7800
      Width           =   6495
   End
   Begin VB.Shape Shape3 
      Height          =   975
      Left            =   240
      Top             =   8280
      Width           =   8415
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total Fare"
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   1000
      TabIndex        =   22
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "total"
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   13920
      TabIndex        =   20
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "senior citizen"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   9360
      TabIndex        =   16
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "children"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   9840
      TabIndex        =   15
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Adult"
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   10200
      TabIndex        =   14
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Travellers and Class:-"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   9120
      TabIndex        =   13
      Top             =   2040
      Width           =   6375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C000&
      BorderWidth     =   6
      Height          =   6015
      Left            =   8880
      Top             =   1560
      Width           =   7815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   5535
      Left            =   240
      Top             =   2640
      Width           =   8415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "journey date"
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   1005
      TabIndex        =   7
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "tax"
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   1000
      TabIndex        =   6
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "fare "
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   1000
      TabIndex        =   5
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "destination"
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   1000
      TabIndex        =   4
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Source "
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   1000
      TabIndex        =   3
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label2 
      ForeColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      ForeColor       =   &H00404080&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "o1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Sub display()
Text1.Text = rs!a
Text2.Text = rs!b
DTPicker1.Value = rs!c
Text3.Text = rs!d
Text4.Text = rs!e
Text5.Text = rs!f
Text6.Text = rs!g
Text7.Text = rs!h
Text8.Text = rs!i
Text9.Text = rs!j
Combo1.Text = rs!k
Combo2.Text = rs!l

End Sub
Private Sub Command3_Click()
 rs.Fields("a").Value = Text1.Text
 rs.Fields("b").Value = Text2.Text
 rs.Fields("c").Value = DTPicker1.Value
 rs.Fields("d").Value = Text3.Text
 rs.Fields("e").Value = Text4.Text
 rs.Fields("f").Value = Text5.Text
 rs.Fields("g").Value = Text6.Text
 rs.Fields("h").Value = Text7.Text
 rs.Fields("i").Value = Text8.Text
 rs.Fields("j").Value = Text9.Text
 rs.Fields("k").Value = Combo1.Text
 rs.Fields("l").Value = Combo2.Text
MsgBox "data has been saved !!", vbInformation, "flight booked "
rs.Update
End Sub

Private Sub Command4_Click()
rs.AddNew
clear
End Sub
'rem we are declaring a sub procedure
Sub clear()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Combo1 = "select flight "
Combo2 = "select cabin class "
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
End Sub

Private Sub Command5_Click()
rs.MoveNext
If Not rs.EOF Then
display
rs.MoveLast
display
End If
End Sub

Private Sub Command6_Click()
rs.MovePrevious
If Not rs.BOF Then
display
rs.MoveLast
display
End If
End Sub

Private Sub Form_Load()
con.Open "provider= microsoft.jet.oledb.4.0;data source=F:\air ticker booking files\Book1.mdb ;Persist security info ="
rs.Open "select * from amber", con, adOpenDynamic, adLockPessimistic
Label2 = Date
Label1 = Time
Combo1.AddItem "f101"
Combo1.AddItem "f102"
Combo1.AddItem "f103"
Combo1.AddItem "f104"
Combo1.AddItem "f105"
Combo2.AddItem "First class"
Combo2.AddItem "Business class"
Combo2.AddItem "Economy class"
Combo2.AddItem "Premium Economy class"
End Sub

Private Sub Combo1_CLICK()
If Combo1 = "f101" Then
Text1 = "mumbai"
Text2 = "patna"
Text3 = 4500
Text4 = 610
Text5 = Val(Text3) + Val(Text4)
ElseIf Combo1 = "f102" Then
Text1 = "BADODARA"
Text2 = "mumbai"
Text3 = 8000
Text4 = 650
Text5 = Val(Text3) + Val(Text4)
ElseIf Combo1 = "f103" Then
Text1 = "varanasi"
Text2 = "mumbai"
Text3 = 5000
Text4 = 650
Text5 = Val(Text3) + Val(Text4)
ElseIf Combo1 = "f104" Then
Text1 = "patna"
Text2 = "ahmedabad"
Text3 = 5500
Text4 = 700
Text5 = Val(Text3) + Val(Text4)
ElseIf Combo1 = "f105" Then
Text1 = "mumbai"
Text2 = "varanasi"
Text3 = 8000
Text4 = 700
Text5 = Val(Text3) + Val(Text4)
End If
End Sub
Private Sub COMMAND1_CLICK()
MsgBox "please provide passenger details ,thank you. "
Form2.Show
o1.Hide
End Sub
Private Sub Command2_Click()
MsgBox "THANK YOU FOR VISITING US"
End
End Sub


Private Sub Text9_Click()
Text9.Text = Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)
End Sub





