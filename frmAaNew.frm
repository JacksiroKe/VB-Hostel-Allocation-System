VERSION 5.00
Begin VB.Form frmAaNew 
   Caption         =   "Add a New Student"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   12
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   11
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Expected Fees:"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   6240
      TabIndex        =   8
      Top             =   360
      Width           =   3015
      Begin VB.Label lblExpectFees 
         Caption         =   "Expected amount"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add a New Staff Member: "
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin VB.TextBox txtExpectedFees 
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Text            =   "fees paid"
         Top             =   2640
         Width           =   4695
      End
      Begin VB.ComboBox cmbDept1 
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "frmAaNew.frx":0000
         Left            =   480
         List            =   "frmAaNew.frx":001C
         TabIndex        =   6
         Text            =   "Choose a Department"
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox txtFullName 
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Text            =   "First and Last Name"
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox txtAdmno 
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Text            =   "Admision no."
         Top             =   1920
         Width           =   4695
      End
      Begin VB.OptionButton optSex 
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1800
         TabIndex        =   3
         Top             =   3480
         Width           =   1215
      End
      Begin VB.OptionButton optSex 
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   3120
         TabIndex        =   2
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "Add this Student"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   1
         Top             =   4320
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   3600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAaNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim sex As String

Private Sub cmbDept1_Click()
    
    If cmbDept1.Text = "Instructor Training" Then
        lblExpectFees.Caption = 24500
    ElseIf cmbDept1.Text = "Mechanical" Then
        lblExpectFees.Caption = 32500
    ElseIf cmbDept1.Text = "Electrical" Then
        lblExpectFees.Caption = 40500
    ElseIf cmbDept1.Text = "Business" Then
        lblExpectFees.Caption = 31500
    ElseIf cmbDept1.Text = "Clothing" Then
        lblExpectFees.Caption = 27500
    ElseIf cmbDept1.Text = "Foods and Beverage" Then
        lblExpectFees.Caption = 35500
    ElseIf cmbDept1.Text = "Computer" Then
        lblExpectFees.Caption = 39500
    ElseIf cmbDept1.Text = "Building & Civil" Then
        lblExpectFees.Caption = 32500
    End If
    lblExpectFees.Visible = True
End Sub

Private Sub optSex_Click(Index As Integer)
    If (Index = 1) Then
        sex = "F"
    Else
        sex = "M"
    End If
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Epamis.mdb;"
    con.Open
    lblExpectFees.Visible = False
End Sub

Private Sub cmdAddNew_Click()
'On Error GoTo ErrorHandler
    If Trim(txtFullName.Text) = "" Or Trim(txtFullName.Text) = "First and Last Name" Or Trim(txtAdmno.Text) = "" Or Trim(txtAdmno.Text) = "Admision no." Then
        MsgBox "Incomplete information provided, Enter all fields to continue", vbCritical, "Validation"
        Exit Sub
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from students", con, adOpenKeyset, adLockOptimistic
    Rs.AddNew
    Rs!fullnames = txtFullName.Text
    Rs!sex = sex
    Rs!feearears = lblExpectFees.Caption - txtExpectedFees.Text
    Rs!admno = txtAdmno.Text
    Rs.Update
    Rs.Close
    MsgBox "New Student was Registered succesfully", vbInformation, App.Title
    Unload Me
    Unload frmAaManage
    frmAaManage.Show
    
 Exit Sub
'ErrorHandler:
'MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub txtAdmno_Click()
    txtAdmno.Text = ""
End Sub

Private Sub txtExpectedFees_Click()
    txtExpectedFees.Text = ""
End Sub

Private Sub txtFullName_Click()
    txtFullName.Text = ""
End Sub
