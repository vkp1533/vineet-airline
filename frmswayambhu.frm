VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frmswayambhu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "swayambhu"
   ClientHeight    =   5205
   ClientLeft      =   1095
   ClientTop       =   390
   ClientWidth     =   11730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   11730
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11730
      TabIndex        =   6
      Top             =   4575
      Width           =   11730
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1213
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "name"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "idnum"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "age"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   4875
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=F:\air ticker booking files\vi.MDB;"
      OLEDBString     =   "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=F:\air ticker booking files\vi.MDB;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select age,idnum,name from swayambhu Order by age"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblLabels 
      Caption         =   "name:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "idnum:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "age:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmswayambhu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

'Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
 ' MsgBox "Data error event hit err:" & Description
'End Sub

'Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal precordset As ADODB.recordset)
  'This will display the current record position for this recordset
 ' datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.recordset.AbsolutePosition)
'End Sub

'Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal precordset As ADODB.recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
 ' Dim bCancel As Boolean

  'Select Case adReason
  'Case adRsnAddNew
  'Case adRsnClose
 ' Case adRsnDelete
  'Case adRsnFirstChange
  'Case adRsnMove
  'Case adRsnRequery
  'Case adRsnResynch
  'Case adRsnUndoAddNew
  'Case adRsnUndoDelete
  'Case adRsnUndoUpdate
  'Case adRsnUpdate
  'End Select

  'If bCancel Then adStatus = adStatusCancel
'End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.recordset.AddNew

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

