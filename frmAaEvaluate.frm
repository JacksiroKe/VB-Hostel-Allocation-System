VERSION 5.00
Begin VB.Form frmAaEvaluate 
   Caption         =   "Evaluate a Member of Staff Now!!!"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14775
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   14775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Evaluate this Staff"
      Height          =   6135
      Left            =   6840
      TabIndex        =   3
      Top             =   120
      Width           =   7575
      Begin VB.OptionButton rate 
         Caption         =   "Poor"
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   15
         Top             =   2800
         Width           =   975
      End
      Begin VB.OptionButton rate 
         Caption         =   "Fair"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   14
         Top             =   2800
         Width           =   1095
      End
      Begin VB.OptionButton rate 
         Caption         =   "Average"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   13
         Top             =   2800
         Width           =   1215
      End
      Begin VB.OptionButton rate 
         Caption         =   "Good"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   12
         Top             =   2800
         Width           =   1095
      End
      Begin VB.OptionButton rate 
         Caption         =   "Excellent"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   2800
         Width           =   1215
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit your Evaluation"
         Height          =   495
         Left            =   720
         TabIndex        =   8
         Top             =   5400
         Width           =   5055
      End
      Begin VB.TextBox txtComment 
         Height          =   855
         Left            =   360
         TabIndex        =   7
         Text            =   "this teacher ...."
         Top             =   4320
         Width           =   6495
      End
      Begin VB.Label Label4 
         Caption         =   "Staff ID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Department:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblStaffid 
         Caption         =   "Staffid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   10
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   7560
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label lblDept 
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label lblName 
         Caption         =   "Staff Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "1. How is this teacher's teaching Method?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   5
         Top             =   2160
         Width           =   6015
      End
      Begin VB.Label Label3 
         Caption         =   "2. In your own words give a recomendation about this teacher."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   4
         Top             =   3600
         Width           =   6015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose a Staff to Evaluate: "
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.ComboBox cmbDept 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   2
         Text            =   "Choose a Department"
         Top             =   480
         Width           =   5415
      End
      Begin VB.ListBox lstStaff 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4110
         Left            =   360
         TabIndex        =   1
         Top             =   1320
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmAaEvaluate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim r1 As Integer, r2 As Integer, r3 As Integer, r4 As Integer, r5 As Integer
Dim rating As String

Private Sub UpdateStaff()
Set Rs = New ADODB.Recordset
    Rs.Open "Select * from staff_members WHERE staffno = '" & lblStaffid.Caption & "'", con, adOpenKeyset, adLockOptimistic
    Rs!rate1 = Rs!rate1 + r1
    Rs!rate2 = Rs!rate2 + r2
    Rs!rate3 = Rs!rate3 + r3
    Rs!rate4 = Rs!rate4 + r4
    Rs!rate5 = Rs!rate5 + r5
    Rs.Update
End Sub

Private Sub NewFeedback()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from feedback", con, adOpenKeyset, adLockOptimistic
    Rs.AddNew
    Rs!rate = rating
    Rs!staffid = lblStaffid.Caption
    Rs!comment = txtComment.Text
    Rs.Update
    
End Sub

Private Sub cmdSubmit_Click()
    'On Error GoTo ErrorHandler
        UpdateStaff
        NewFeedback
    Rs.Close
    MsgBox "You Have evaluated succesfully!", vbInformation, App.Title
    frmAaEvaluate.Show
    Unload Me
Exit Sub
'ErrorHandler:
 'MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\pucspes.mdb;"
    con.Open
    Load_Latest
End Sub

Private Sub Load_Latest()
lstStaff.Clear
On Error GoTo ErrorHandler
 Set Rs = New ADODB.Recordset
    Rs.Open "Select * from staff_members ORDER BY staffid ASC", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        lstStaff.AddItem Rs!staffname
        Rs.MoveNext
    Loop
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub


Private Sub lstStaff_Click()
On Error GoTo ErrorHandler
Set Rs = New ADODB.Recordset
Rs.Open "Select * from staff_members WHERE staffname = '" & lstStaff.Text & "'", con, adOpenKeyset, adLockOptimistic
lblName = Rs!staffname
lblDept = Rs!department
lblStaffid = Rs!staffno
Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub rate_Click(Index As Integer)
    If (Index = 0) Then
        r1 = 1
        r2 = 0
        r3 = 0
        r4 = 0
        r5 = 0
        rating = "rate5"
    ElseIf (Index = 1) Then
        r1 = 0
        r2 = 1
        r3 = 0
        r4 = 0
        r5 = 0
        rating = "rate4"
    ElseIf (Index = 2) Then
        r1 = 0
        r2 = 0
        r3 = 1
        r4 = 0
        r5 = 0
        rating = "rate3"
    ElseIf (Index = 3) Then
        r1 = 0
        r2 = 0
        r3 = 0
        r4 = 1
        r5 = 0
        rating = "rate2"
    ElseIf (Index = 4) Then
        r1 = 0
        r2 = 0
        r3 = 0
        r4 = 0
        r5 = 1
        rating = "rate1"
    End If
End Sub

Private Sub txtComment_Click()
    txtComment.Text = ""
End Sub
