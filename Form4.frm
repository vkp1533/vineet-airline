VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form4 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form4"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Amita"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   12495
   ScaleWidth      =   22920
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Add More"
      BeginProperty Font 
         Name            =   "Amita"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   8400
      TabIndex        =   20
      Top             =   9360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Round Trip Ticket Booking "
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11415
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   16335
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   735
         Left            =   11400
         TabIndex        =   33
         Top             =   5880
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1296
         _Version        =   393216
         Format          =   122355713
         CurrentDate     =   45050
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   915
         Left            =   11400
         TabIndex        =   31
         Top             =   3840
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   735
         Left            =   11400
         TabIndex        =   29
         Top             =   4920
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1296
         _Version        =   393216
         CalendarBackColor=   -2147483646
         Format          =   122355713
         CurrentDate     =   45049
      End
      Begin VB.Timer Timer1 
         Left            =   120
         Top             =   1920
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Back"
         Height          =   1215
         Left            =   11640
         TabIndex        =   25
         Top             =   9840
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Confirm "
         Height          =   1215
         Left            =   11640
         TabIndex        =   24
         Top             =   8040
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFC0&
         Height          =   915
         Left            =   11400
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FFFFC0&
         Height          =   915
         Left            =   11400
         TabIndex        =   9
         Top             =   1680
         Width           =   2775
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFFFC0&
         Height          =   915
         Left            =   11400
         TabIndex        =   8
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Passenger's Details"
         ForeColor       =   &H00808000&
         Height          =   4455
         Left            =   480
         TabIndex        =   1
         Top             =   6840
         Width           =   9495
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Amita"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4560
            TabIndex        =   22
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Save"
            Height          =   735
            Left            =   7320
            TabIndex        =   19
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Amita"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3360
            TabIndex        =   4
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Amita"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3360
            TabIndex        =   3
            Top             =   2520
            Width           =   2535
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Amita"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   3360
            TabIndex        =   2
            Top             =   3480
            Width           =   2535
         End
         Begin VB.Label Label8 
            Caption         =   "TOTAL PASSENGERS :-"
            Height          =   735
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   4455
         End
         Begin VB.Label Label5 
            Caption         =   "Host Name"
            Height          =   735
            Left            =   120
            TabIndex        =   7
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label Label6 
            Caption         =   "age"
            Height          =   735
            Left            =   120
            TabIndex        =   6
            Top             =   2520
            Width           =   2655
         End
         Begin VB.Label Label7 
            Caption         =   "id number"
            Height          =   735
            Left            =   120
            TabIndex        =   5
            Top             =   3480
            Width           =   2655
         End
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OF RETURN:-"
         Height          =   735
         Left            =   6720
         TabIndex        =   32
         Top             =   5880
         Width           =   4455
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL FARE:-"
         Height          =   855
         Left            =   6840
         TabIndex        =   30
         Top             =   3840
         Width           =   4095
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OF DEPARTURE:-"
         Height          =   855
         Left            =   6720
         TabIndex        =   28
         Top             =   4920
         Width           =   4695
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   480
         TabIndex        =   27
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   600
         TabIndex        =   26
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Round Trip Ticket Booking :-"
         BeginProperty Font 
            Name            =   "Amita"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   23
         Top             =   -120
         Width           =   8175
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "preferred  Airways:-"
         Height          =   855
         Left            =   6840
         TabIndex        =   18
         Top             =   2760
         Width           =   4575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Destination :-"
         Height          =   855
         Left            =   6840
         TabIndex        =   17
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Origin :-"
         ForeColor       =   &H80000017&
         Height          =   855
         Left            =   6840
         TabIndex        =   15
         Top             =   720
         Width           =   3975
      End
      Begin VB.Image Image1 
         Height          =   15000
         Left            =   -5880
         MousePointer    =   1  'Arrow
         Picture         =   "Form4.frx":1ED62
         Top             =   -3600
         Width           =   22500
      End
      Begin VB.Label Label1 
         Caption         =   "Select Origin"
         Height          =   855
         Left            =   720
         TabIndex        =   14
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Select Destination "
         Height          =   855
         Left            =   720
         TabIndex        =   13
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Label3 
         Caption         =   "Select preffered airyways "
         BeginProperty Font 
            Name            =   "Amita"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   720
         TabIndex        =   12
         Top             =   3120
         Width           =   3855
      End
      Begin VB.Label Label4 
         Height          =   495
         Left            =   960
         TabIndex        =   11
         Top             =   5160
         Width           =   2655
      End
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   495
      Left            =   10920
      TabIndex        =   16
      Top             =   6000
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
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
Private Sub COMMAND1_CLICK()
 rs.Fields("ORIGIN").Value = Combo1.Text
 rs.Fields("DESTINATION").Value = Combo2.Text
 rs.Fields("AIRLINE").Value = Combo3.Text
 rs.Fields("DATE1").Value = DTPicker1.Value
 rs.Fields("DATE2").Value = DTPicker2.Value
 rs.Fields("NAME").Value = Text1.Text
 rs.Fields("AGE").Value = Text2.Text
 rs.Fields("IDN").Value = Text3.Text
 rs.Fields("TPASSENGER").Value = Text4.Text
MsgBox "data has been saved !!", vbInformation, "flight booked "
rs.Update
End Sub
Private Sub Command3_Click()
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
Private Sub Command4_Click()
MsgBox "THANK YOU ,FOLLOW NEXT PAGE "
Form4.Hide
Form5.Show
End Sub

Private Sub Command5_Click()
Form4.Hide
End Sub

Private Sub Form_Load()
con.Open "provider= microsoft.jet.oledb.4.0;data source=F:\air ticker booking files\Book1.mdb ;Persist security info ="
rs.Open "select * from TRINETRADHARI", con, adOpenDynamic, adLockPessimistic
Label14 = Date
Label15 = Time
Combo1.AddItem "VARANASI"
Combo1.AddItem "MUMBAI"
Combo1.AddItem "PATNA"
Combo1.AddItem "SURAT"
Combo1.AddItem "AGRA"
Combo1.AddItem "SRIHARIKOTA"
Combo2.AddItem "SRIHARIKOTA"
Combo2.AddItem "MUMBAI"
Combo2.AddItem "CHATTISGARH"
Combo2.AddItem "PUNJAB"
Combo2.AddItem "GORAKHPUR"
Combo3.AddItem "QATAR AIRWAYS"
Combo3.AddItem "AIR INDIA"
Combo3.AddItem "INDIGO"
Combo3.AddItem "SPICEJET"
Combo3.AddItem "JET AIRWAYS"
End Sub

Private Sub COMBO3_CLICK()
If Combo1 = "VARANASI" And Combo2 = "MUMBAI" And Combo3 = "AIR INDIA" Then
Text5 = 20000
ElseIf Combo1 = "VARANASI" And Combo2 = "PUNJAB" And Combo3 = "AIR INIDA" Then
Text5 = 21000
ElseIf Combo1 = "VARANASI" And Combo2 = "GORAKHPUR" And Combo3 = "AIR INIDA" Then
Text5 = 6122
ElseIf Combo1 = "VARANASI" And Combo2 = "SRIHARIKOTA" And Combo3 = "AIR INIDA" Then
Text5 = 45000
ElseIf Combo1 = "VARANASI" And Combo2 = "CHATTISGARH" And Combo3 = "AIR INIDA" Then
Text5 = 32000
ElseIf Combo1 = "PATNA" And Combo2 = "PUNJAB" And Combo3 = "AIR INIDA" Then
Text5 = 67000
ElseIf Combo1 = "PATNA" And Combo2 = "GORAKHPUR" And Combo3 = "AIR INIDA" Then
Text5 = 11222
ElseIf Combo1 = "PATNA" And Combo2 = "SRIHARIKOTA" And Combo3 = "AIR INIDA" Then
Text5 = 64220
ElseIf Combo1 = "PATNA" And Combo2 = "CHATTISGARH" And Combo3 = "AIR INIDA" Then
Text5 = 55444
ElseIf Combo1 = "PATNA" And Combo2 = "MUMBAI" And Combo3 = "AIR INIDA" Then
Text5 = 54433
ElseIf Combo1 = "MUMBAI" And Combo2 = "PUNJAB" And Combo3 = "AIR INIDA" Then
Text5 = 56666
ElseIf Combo1 = "MUMBAI" And Combo2 = "GORAKHPUR" And Combo3 = "AIR INIDA" Then
Text5 = 54343
ElseIf Combo1 = "MUMBAI" And Combo2 = "SRIHARIKOTA" And Combo3 = "AIR INIDA" Then
Text5 = 54544

ElseIf Combo1 = "MUMBAI" And Combo2 = "CHATTISGARH" And Combo3 = "AIR INIDA" Then
Text5 = 14566
ElseIf Combo1 = "MUMBAI" And Combo2 = "MUMBAI" And Combo3 = "AIR INIDA" Then
Text5 = 0
ElseIf Combo1 = "SURAT" And Combo2 = "PUNJAB" And Combo3 = "AIR INIDA" Then
Text5 = 18000
ElseIf Combo1 = "SURAT" And Combo2 = "GORAKHPUR" And Combo3 = "AIR INIDA" Then
Text5 = 58000
ElseIf Combo1 = "SURAT" And Combo2 = "SRIHARIKOTA" And Combo3 = "AIR INIDA" Then
Text5 = 45665

ElseIf Combo1 = "SURAT" And Combo2 = "CHATTISGARH" And Combo3 = "AIR INIDA" Then
Text5 = 43333
ElseIf Combo1 = "SURAT" And Combo2 = "MUMBAI" And Combo3 = "AIR INIDA" Then
Text5 = 45555
ElseIf Combo1 = "AGRA" And Combo2 = "PUNJAB" And Combo3 = "AIR INIDA" Then
Text5 = 60000
ElseIf Combo1 = "AGRA" And Combo2 = "GORAKHPUR" And Combo3 = "AIR INIDA" Then
Text5 = 35000
ElseIf Combo1 = "AGRA" And Combo2 = "SRIHARIKOTA" And Combo3 = "AIR INIDA" Then
Text5 = 167867

ElseIf Combo1 = "AGRA" And Combo2 = "CHATTISGARH" And Combo3 = "AIR INIDA" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "MUMBAI" And Combo3 = "AIR INIDA" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "PUNJAB" And Combo3 = "AIR INIDA" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "GORAKHPUR" And Combo3 = "AIR INIDA" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "SRIHARIKOTA" And Combo3 = "AIR INIDA" Then
Text5 = X

ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "CHATTISGARH" And Combo3 = "AIR INIDA" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "MUMBAI" And Combo3 = "AIR INIDA" Then
Text5 = X
ElseIf Combo1 = "VARANASI" And Combo2 = "MUMBAI" And Combo3 = "AIR INDIA" Then
Text5 = X
ElseIf Combo1 = "VARANASI" And Combo2 = "PUNJAB" And Combo3 = "SPICEJET" Then
Text5 = 23456
ElseIf Combo1 = "VARANASI" And Combo2 = "GORAKHPUR" And Combo3 = "SPICEJET" Then
Text5 = 32456
ElseIf Combo1 = "VARANASI" And Combo2 = "SRIHARIKOTA" And Combo3 = "SPICEJET" Then
Text5 = 140000
ElseIf Combo1 = "VARANASI" And Combo2 = "CHATTISGARH" And Combo3 = "SPICEJET" Then
Text5 = 40000
ElseIf Combo1 = "PATNA" And Combo2 = "PUNJAB" And Combo3 = "SPICEJET" Then
Text5 = 54332
ElseIf Combo1 = "PATNA" And Combo2 = "GORAKHPUR" And Combo3 = "SPICEJET" Then
Text5 = 56676
ElseIf Combo1 = "PATNA" And Combo2 = "SRIHARIKOTA" And Combo3 = "SPICEJET" Then
Text5 = X

ElseIf Combo1 = "PATNA" And Combo2 = "CHATTISGARH" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "PATNA" And Combo2 = "MUMBAI" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "MUMBAI" And Combo2 = "PUNJAB" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "MUMBAI" And Combo2 = "GORAKHPUR" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "MUMBAI" And Combo2 = "SRIHARIKOTA" And Combo3 = "SPICEJET" Then
Text5 = X

ElseIf Combo1 = "MUMBAI" And Combo2 = "CHATTISGARH" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "MUMBAI" And Combo2 = "MUMBAI" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "SURAT" And Combo2 = "PUNJAB" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "SURAT" And Combo2 = "GORAKHPUR" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "SURAT" And Combo2 = "SRIHARIKOTA" And Combo3 = "SPICEJET" Then
Text5 = X

ElseIf Combo1 = "SURAT" And Combo2 = "CHATTISGARH" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "SURAT" And Combo2 = "MUMBAI" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "PUNJAB" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "GORAKHPUR" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "SRIHARIKOTA" And Combo3 = "SPICEJET" Then
Text5 = X

ElseIf Combo1 = "AGRA" And Combo2 = "CHATTISGARH" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "MUMBAI" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "PUNJAB" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "GORAKHPUR" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "SRIHARIKOTA" And Combo3 = "SPICEJET" Then
Text5 = X

ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "CHATTISGARH" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "MUMBAI" And Combo3 = "SPICEJET" Then
Text5 = X
ElseIf Combo1 = "VARANASI" And Combo2 = "MUMBAI" And Combo3 = "AIR INDIA" Then
Text5 = X
ElseIf Combo1 = "VARANASI" And Combo2 = "PUNJAB" And Combo3 = "QATAR AIRWAYS" Then
Text5 = 45678
ElseIf Combo1 = "VARANASI" And Combo2 = "GORAKHPUR" And Combo3 = "QATAR AIRWAYS" Then
Text5 = 10000
ElseIf Combo1 = "VARANASI" And Combo2 = "SRIHARIKOTA" And Combo3 = "QATAR AIRWAYS" Then
Text5 = 200000
ElseIf Combo1 = "VARANASI" And Combo2 = "CHATTISGARH" And Combo3 = "QATAR AIRWAYS" Then
Text5 = 53444
ElseIf Combo1 = "PATNA" And Combo2 = "PUNJAB" And Combo3 = "QATAR AIRWAYS" Then
Text5 = 87654
ElseIf Combo1 = "PATNA" And Combo2 = "GORAKHPUR" And Combo3 = "QATAR AIRWAYS" Then
Text5 = 80000
ElseIf Combo1 = "PATNA" And Combo2 = "SRIHARIKOTA" And Combo3 = "QATAR AIRWAYS" Then
Text5 = 254321

ElseIf Combo1 = "PATNA" And Combo2 = "CHATTISGARH" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "PATNA" And Combo2 = "MUMBAI" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "MUMBAI" And Combo2 = "PUNJAB" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "MUMBAI" And Combo2 = "GORAKHPUR" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "MUMBAI" And Combo2 = "SRIHARIKOTA" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X

ElseIf Combo1 = "MUMBAI" And Combo2 = "CHATTISGARH" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "MUMBAI" And Combo2 = "MUMBAI" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SURAT" And Combo2 = "PUNJAB" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SURAT" And Combo2 = "GORAKHPUR" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SURAT" And Combo2 = "SRIHARIKOTA" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X

ElseIf Combo1 = "SURAT" And Combo2 = "CHATTISGARH" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SURAT" And Combo2 = "MUMBAI" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "PUNJAB" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "GORAKHPUR" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "SRIHARIKOTA" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X

ElseIf Combo1 = "AGRA" And Combo2 = "CHATTISGARH" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "MUMBAI" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "PUNJAB" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "GORAKHPUR" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "SRIHARIKOTA" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X

ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "CHATTISGARH" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "MUMBAI" And Combo3 = "QATAR AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "VARANASI" And Combo2 = "MUMBAI" And Combo3 = "AIR INDIA" Then
Text5 = 32000
ElseIf Combo1 = "VARANASI" And Combo2 = "PUNJAB" And Combo3 = "INDIGO" Then
Text5 = 54000
ElseIf Combo1 = "VARANASI" And Combo2 = "GORAKHPUR" And Combo3 = "INDIGO" Then
Text5 = 45000
ElseIf Combo1 = "VARANASI" And Combo2 = "SRIHARIKOTA" And Combo3 = "INDIGO" Then
Text5 = 120000
ElseIf Combo1 = "VARANASI" And Combo2 = "CHATTISGARH" And Combo3 = "INDIGO" Then
Text5 = 34567
ElseIf Combo1 = "PATNA" And Combo2 = "PUNJAB" And Combo3 = "INDIGO" Then
Text5 = 43215
ElseIf Combo1 = "PATNA" And Combo2 = "GORAKHPUR" And Combo3 = "INDIGO" Then
Text5 = 45678
ElseIf Combo1 = "PATNA" And Combo2 = "SRIHARIKOTA" And Combo3 = "INDIGO" Then
Text5 = 156780

ElseIf Combo1 = "PATNA" And Combo2 = "CHATTISGARH" And Combo3 = "INDIGO" Then
Text5 = 45734
ElseIf Combo1 = "PATNA" And Combo2 = "MUMBAI" And Combo3 = "INDIGO" Then
Text5 = 34567
ElseIf Combo1 = "MUMBAI" And Combo2 = "PUNJAB" And Combo3 = "INDIGO" Then
Text5 = 65432
ElseIf Combo1 = "MUMBAI" And Combo2 = "GORAKHPUR" And Combo3 = "INDIGO" Then
Text5 = 34567
ElseIf Combo1 = "MUMBAI" And Combo2 = "SRIHARIKOTA" And Combo3 = "INDIGO" Then
Text5 = 98765

ElseIf Combo1 = "MUMBAI" And Combo2 = "CHATTISGARH" And Combo3 = "INDIGO" Then
Text5 = 87654
ElseIf Combo1 = "MUMBAI" And Combo2 = "MUMBAI" And Combo3 = "INDIGO" Then
Text5 = 65432
ElseIf Combo1 = "SURAT" And Combo2 = "PUNJAB" And Combo3 = "INDIGO" Then
Text5 = 76543
ElseIf Combo1 = "SURAT" And Combo2 = "GORAKHPUR" And Combo3 = "INDIGO" Then
Text5 = 98765
ElseIf Combo1 = "SURAT" And Combo2 = "SRIHARIKOTA" And Combo3 = "INDIGO" Then
Text5 = X

ElseIf Combo1 = "SURAT" And Combo2 = "CHATTISGARH" And Combo3 = "INDIGO" Then
Text5 = X
ElseIf Combo1 = "SURAT" And Combo2 = "MUMBAI" And Combo3 = "INDIGO" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "PUNJAB" And Combo3 = "INDIGO" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "GORAKHPUR" And Combo3 = "INDIGO" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "SRIHARIKOTA" And Combo3 = "INDIGO" Then
Text5 = X

ElseIf Combo1 = "AGRA" And Combo2 = "CHATTISGARH" And Combo3 = "INDIGO" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "MUMBAI" And Combo3 = "INDIGO" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "PUNJAB" And Combo3 = "INDIGO" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "GORAKHPUR" And Combo3 = "INDIGO" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "SRIHARIKOTA" And Combo3 = "INDIGO" Then
Text5 = X

ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "CHATTISGARH" And Combo3 = "INDIGO" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "MUMBAI" And Combo3 = "INDIGO" Then
Text5 = X
ElseIf Combo1 = "VARANASI" And Combo2 = "MUMBAI" And Combo3 = "AIR INDIA" Then
Text5 = X
ElseIf Combo1 = "VARANASI" And Combo2 = "PUNJAB" And Combo3 = "JET AIRWAYS" Then
Text5 = 143222
ElseIf Combo1 = "VARANASI" And Combo2 = "GORAKHPUR" And Combo3 = "JET AIRWAYS" Then
Text5 = 12233
ElseIf Combo1 = "VARANASI" And Combo2 = "SRIHARIKOTA" And Combo3 = "JET AIRWAYS" Then
Text5 = 123345
ElseIf Combo1 = "VARANASI" And Combo2 = "CHATTISGARH" And Combo3 = "JET AIRWAYS" Then
Text5 = 123456
ElseIf Combo1 = "PATNA" And Combo2 = "PUNJAB" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "PATNA" And Combo2 = "GORAKHPUR" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "PATNA" And Combo2 = "SRIHARIKOTA" And Combo3 = "JET AIRWAYS" Then
Text5 = X

ElseIf Combo1 = "PATNA" And Combo2 = "CHATTISGARH" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "PATNA" And Combo2 = "MUMBAI" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "MUMBAI" And Combo2 = "PUNJAB" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "MUMBAI" And Combo2 = "GORAKHPUR" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "MUMBAI" And Combo2 = "SRIHARIKOTA" And Combo3 = "JET AIRWAYS" Then
Text5 = X

ElseIf Combo1 = "MUMBAI" And Combo2 = "CHATTISGARH" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "MUMBAI" And Combo2 = "MUMBAI" And Combo3 = "JET AIRWAYS" Then
Text5 = 0
ElseIf Combo1 = "SURAT" And Combo2 = "PUNJAB" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SURAT" And Combo2 = "GORAKHPUR" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SURAT" And Combo2 = "SRIHARIKOTA" And Combo3 = "JET AIRWAYS" Then
Text5 = X

ElseIf Combo1 = "SURAT" And Combo2 = "CHATTISGARH" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SURAT" And Combo2 = "MUMBAI" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "PUNJAB" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "GORAKHPUR" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "SRIHARIKOTA" And Combo3 = "JET AIRWAYS" Then
Text5 = X

ElseIf Combo1 = "AGRA" And Combo2 = "CHATTISGARH" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "AGRA" And Combo2 = "MUMBAI" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "PUNJAB" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "GORAKHPUR" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "SRIHARIKOTA" And Combo3 = "JET AIRWAYS" Then
Text5 = 0
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "CHATTISGARH" And Combo3 = "JET AIRWAYS" Then
Text5 = X
ElseIf Combo1 = "SRIHARIKOTA" And Combo2 = "MUMBAI" And Combo3 = "JET AIRWAYS" Then
Text5 = X



End If
End Sub




