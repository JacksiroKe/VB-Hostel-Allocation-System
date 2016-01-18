VERSION 5.00
Begin VB.Form frmAaManage 
   Caption         =   "Eldoret Polytechnic Accomodation MIS"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   21.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbCriteria 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      ItemData        =   "frmAaManage.frx":0000
      Left            =   9840
      List            =   "frmAaManage.frx":0013
      TabIndex        =   14
      Text            =   "Choose a Criteria"
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Frame fraShow 
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   9360
      TabIndex        =   9
      Top             =   2640
      Width           =   5535
      Begin VB.Label lblFisrt 
         Caption         =   "First Label"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label lblSecond 
         Caption         =   "Second Label"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   5055
      End
      Begin VB.Label lblThird 
         Caption         =   "Third Label"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   5175
      End
      Begin VB.Label lblFourth 
         Caption         =   "Fourth Label"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
      Begin VB.CommandButton cmdLogOut 
         Caption         =   "Log Out"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   3840
         Width           =   2055
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Allocation"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton cmdAllocate 
         Caption         =   "Allocate  Room"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton cmdNewStude 
         Caption         =   "New Student"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   6495
      Begin VB.ListBox lstStudents 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5715
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Display Criteria"
      Height          =   615
      Left            =   10080
      TabIndex        =   8
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Current Enrolment of Students"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "frmAaManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim sstatus As String

Dim sex As String
    
Private Sub optSex_Click(Index As Integer)
    If (Index = 1) Then
        sex = "female"
    Else
        sex = "male"
    End If
End Sub

Private Sub cmbCriteria_Click()
    If cmbCriteria.Text = "Show All Students" Then
        Load_AllStudents
    ElseIf cmbCriteria.Text = "Show Boarding Students" Then
        Load_AllBStudents
    ElseIf cmbCriteria.Text = "Show Commutting Students" Then
        Load_AllCStudents
    ElseIf cmbCriteria.Text = "Show All Rooms" Then
        Load_AllRooms
    ElseIf cmbCriteria.Text = "Show All Vacant Rooms" Then
        LoadAllVRooms
    End If
    
End Sub

Private Sub cmdAllocate_Click()
    If lstStudents.Text = "" Or cmbCriteria.Text = "Show All Rooms" Or cmbCriteria.Text = "Show All Vacant Rooms" Then
        MsgBox "Please click on a student Name first to proceed!", vbCritical, App.Title
    Exit Sub
    End If
    frmAaAllocate.Show
End Sub

Private Sub cmdClear_Click()
    If lstStudents.Text = "" Or cmbCriteria.Text = "Show All Rooms" Or cmbCriteria.Text = "Show All Vacant Rooms" Then
        MsgBox "Please click on a student Name first to proceed!", vbCritical, App.Title
    Exit Sub
    End If
    ChangeStatus
    MsgBox lstStudents.Text & " has been cleared from the room!", vbInformation, App.Title
End Sub

Private Sub ChangeStatus()
 On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from students where FullNames = '" & lstStudents.Text & "'", con, adOpenKeyset, adLockOptimistic
    Rs!Status = "commuter"
    Rs.Update
    Rs.Close
 Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
    
End Sub

Private Sub cmdLogOut_Click()
    frmAaLogin.Show
    Unload Me
End Sub

Private Sub cmdNewStude_Click()
    frmAaNew.Show
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Epamis.mdb;"
    con.Open
    Load_AllStudents
    fraShow.Visible = False
End Sub

Private Sub Load_AllStudents()
lstStudents.Clear
Dim str As String
On Error GoTo ErrorHandler
 Set Rs = New ADODB.Recordset
    Rs.Open "Select * from students ORDER BY sid ASC", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        str = Rs!fullnames & " (" & Rs!sex & ")"
        lstStudents.AddItem Rs!fullnames
        Rs.MoveNext
    Loop
    Rs.Close
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub Load_AllBStudents()
lstStudents.Clear
Dim str As String
Dim sstatus As String
sstatus = "boarder"
On Error GoTo ErrorHandler
 Set Rs = New ADODB.Recordset
    Rs.Open "Select * from students WHERE status= '" & sstatus & "'", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        str = Rs!fullnames & " (" & Rs!sex & ")"
        lstStudents.AddItem Rs!fullnames
        Rs.MoveNext
    Loop
    Rs.Close
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub Load_AllCStudents()
lstStudents.Clear
Dim str As String
sstatus = "commuter"
On Error GoTo ErrorHandler
 Set Rs = New ADODB.Recordset
    Rs.Open "Select * from students WHERE status= '" & sstatus & "'", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        str = Rs!fullnames & " (" & Rs!sex & ")"
        lstStudents.AddItem Rs!fullnames
        Rs.MoveNext
    Loop
    Rs.Close
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub


Private Sub Load_AllRooms()
lstStudents.Clear
Dim str As String
On Error GoTo ErrorHandler
 Set Rs = New ADODB.Recordset
    Rs.Open "Select * from hostels", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        str = Rs!Hostel & " - Room " & Rs!room
        lstStudents.AddItem str
        Rs.MoveNext
    Loop
    Rs.Close
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub LoadAllVRooms()
lstStudents.Clear
Dim str As String
On Error GoTo ErrorHandler
Dim studA As String, studB As String, studC As String, studD As String
studA = "Null"
studB = "Null"
studC = "Null"
studD = "Null"

 Set Rs = New ADODB.Recordset
    Rs.Open "Select * from hostels WHERE StudentA = '" & studA & "' Or StudentB = '" & studB & "' or StudentC = '" & studC & "' or StudentD = '" & studD & "'", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        str = Rs!Hostel & " - Room " & Rs!room
        lstStudents.AddItem str
        Rs.MoveNext
    Loop
    Rs.Close
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub


Private Sub lstStudents_Click()
fraShow.Visible = True
If cmbCriteria.Text = "Show All Students" Then
        StudentDetails
    ElseIf cmbCriteria.Text = "Show Boarding Students" Then
        StudentDetails
    ElseIf cmbCriteria.Text = "Show Commutting Students" Then
        StudentDetails
    ElseIf cmbCriteria.Text = "Show All Rooms" Then
        RoomDetails
    ElseIf cmbCriteria.Text = "Show All Vacant Rooms" Then
        RoomDetails
    Else: StudentDetails
    End If
End Sub

Private Sub StudentDetails()
Dim StudentStatus As String
On Error GoTo ErrorHandler
Set Rs = New ADODB.Recordset
Rs.Open "Select * from students WHERE FullNames = '" & lstStudents.Text & "'", con, adOpenKeyset, adLockOptimistic
fraShow.Caption = "Student Details:"
lblFisrt.Caption = "Adm. No: " & Rs!admno
lblSecond.Caption = "Fee Arears: " & Rs!FeeArears
lblThird.Caption = "Status: " & Rs!Status
StudentStatus = Rs!Status
Rs.Close
StudentRoomDetails
If StudentStatus = "commuter" Then
    lblFourth.Caption = ""
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub StudentRoomDetails()
Set Rs = New ADODB.Recordset
Rs.Open "Select * from hostels WHERE StudentA='" & lstStudents.Text & "' or StudentB='" & lstStudents.Text & "' or StudentC='" & lstStudents.Text & "' or StudentD='" & lstStudents.Text & "'", con, adOpenKeyset, adLockOptimistic
Dim studA As String, studB As String, studC As String, studD As String
studA = Rs!studentA
studB = Rs!studentB
studC = Rs!studentC
studD = Rs!studentD
If studA = "Null" Or studB = "Null" Or studC = "Null" Or studD = "Null" Then
    lblFourth.Caption = ""
End If

If studA = "Null" Then
    lblFourth.Caption = "Room: " & Rs!room
ElseIf Not (studB = "Null") Then
    lblFourth.Caption = "Room: " & Rs!room
ElseIf Not (studC = "Null") Then
    lblFourth.Caption = "Room: " & Rs!room
ElseIf Not (studD = "Null") Then
    lblFourth.Caption = "Room: " & Rs!room
End If
Rs.Close
Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub RoomDetails()
Dim roomid As String
roomid = Trim(Right(lstStudents.Text, 3))

Set Rs = New ADODB.Recordset
Rs.Open "Select * from hostels WHERE room = '" & roomid & "'", con, adOpenKeyset, adLockOptimistic
Dim studA As String, studB As String, studC As String, studD As String
studA = Rs!studentA
studB = Rs!studentB
studC = Rs!studentC
studD = Rs!studentD
fraShow.Caption = "Room " & roomid & ":"

If studA = "Null" Then
    lblFisrt.Caption = "Student A: "
Else
    lblFisrt.Caption = "Student A: " & studA
End If

If studB = "Null" Then
    lblSecond.Caption = "Student B: "
Else
    lblSecond.Caption = "Student B: " & studB
End If

If studC = "Null" Then
    lblThird.Caption = "Student C: "
Else
    lblThird.Caption = "Student C: " & studC
End If

If studD = "Null" Then
    lblFourth.Caption = "Student D: "
Else
    lblFourth.Caption = "Student D: " & studD
End If
Rs.Close
Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub
