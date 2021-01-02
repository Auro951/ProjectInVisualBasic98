VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc StudentRecord 
      Height          =   3135
      Left            =   10680
      Top             =   2040
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   5530
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB98\student database full\StudentRecord.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB98\student database full\StudentRecord.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "StudentInfo"
      Caption         =   "StudentInfo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton clearbtn 
      Caption         =   "Clear"
      Height          =   615
      Left            =   3480
      TabIndex        =   19
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "Delete"
      Height          =   615
      Left            =   8160
      TabIndex        =   18
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton update 
      Caption         =   "Update"
      Height          =   615
      Left            =   8160
      TabIndex        =   17
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton addbtn 
      Caption         =   "Add"
      Height          =   615
      Left            =   8160
      TabIndex        =   16
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton lastbtn 
      Caption         =   "last"
      Height          =   615
      Left            =   8160
      TabIndex        =   15
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   "next"
      Height          =   615
      Left            =   8160
      TabIndex        =   14
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton previousbtn 
      Caption         =   "previous"
      Height          =   615
      Left            =   8160
      TabIndex        =   13
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton firstbtn 
      Caption         =   "first"
      Height          =   615
      Left            =   8160
      TabIndex        =   12
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox txtphone 
      DataField       =   "phoneno"
      DataSource      =   "StudentRecord"
      Height          =   615
      Left            =   3600
      TabIndex        =   11
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox txtemail 
      DataField       =   "emailid"
      DataSource      =   "StudentRecord"
      Height          =   615
      Left            =   3600
      TabIndex        =   10
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox txtaddress 
      DataField       =   "address"
      DataSource      =   "StudentRecord"
      Height          =   615
      Left            =   3600
      TabIndex        =   9
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txtclass 
      DataField       =   "stream"
      DataSource      =   "StudentRecord"
      Height          =   615
      Left            =   3600
      TabIndex        =   8
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox txtname 
      DataField       =   "studentname"
      DataSource      =   "StudentRecord"
      Height          =   615
      Left            =   3600
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtroll 
      DataField       =   "rollno"
      DataSource      =   "StudentRecord"
      Height          =   615
      Left            =   3600
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "      Student Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      TabIndex        =   20
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label phoneno 
      Caption         =   "Phone No."
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label stream 
      Caption         =   "Stream"
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label address 
      Caption         =   "Address"
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label emailid 
      Caption         =   "Email ID"
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label studentname 
      Caption         =   "Student's Name"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label rillno 
      Caption         =   "Roll No."
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub firstbtn_Click()
StudentRecord.Recordset.MoveFirst
End Sub
Private Sub lastbtn_Click()
StudentRecord.Recordset.MoveLast
End Sub
Private Sub nextbtn_Click()
StudentRecord.Recordset.MoveNext
End Sub
Private Sub previousbtn_Click()
StudentRecord.Recordset.MovePrevious
End Sub
Private Sub addbtn_Click()
StudentRecord.Recordset.AddNew
End Sub
Private Sub clearbtn_Click()
txtroll.Text = " "
txtname.Text = " "
txtclass.Text = " "
txtaddress.Text = " "
txtemail.Text = " "
txtphone.Text = " "
End Sub
Private Sub deletebtn_Click()
confirmation = MsgBox("Do you want to delete this record", vbYesNo + vbCritical, "Delete Record Confirmation")
If confirmation = vbYes Then
StudentRecord.Recordset.Delete
MsgBox "Record Has Been Deleted Successfully", vbInformation, "message"
Else
MsgBox "Record Not Deleted", vbInformation, "message"
End If
End Sub
Private Sub update_Click()
StudentRecord.Recordset.Fields("rollno") = txtroll.Text
StudentRecord.Recordset.Fields("studentname") = txtname.Text
StudentRecord.Recordset.Fields("stream") = txtclass.Text
StudentRecord.Recordset.Fields("address") = txtaddress.Text
StudentRecord.Recordset.Fields("emailid") = txtemail.Text
StudentRecord.Recordset.Fields("phoneno") = txtphone.Text
StudentRecord.Recordset.update
MsgBox "data is saved successfully", vbInformation, "message"
End Sub
