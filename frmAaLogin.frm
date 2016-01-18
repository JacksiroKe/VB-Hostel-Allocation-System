VERSION 5.00
Begin VB.Form frmAaLogin 
   Caption         =   "Eldoret Polytechnic Accomodation MIS"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8790
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H8000000B&
      Caption         =   "Login"
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   3120
      Width           =   4335
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00C0FFFF&
      Height          =   555
      Left            =   4320
      TabIndex        =   2
      Text            =   "pass code"
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00C0FFFF&
      Height          =   555
      Left            =   4320
      TabIndex        =   1
      Text            =   "user code"
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter your Pass code:"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter your user code:"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Login to your Account"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmAaLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim accountype As String
Dim username As String
Dim password As String
Dim login As Boolean


Private Sub cmdLogin_Click()
If txtUsername.Text = "" Or txtUsername.Text = "user code" Or txtPassword.Text = "" Or txtPassword.Text = "pass code" Then
    MsgBox "You must enter your details properly", vbCritical, "Eldoret Polytechnic Accomodation MIS"
Else
    frmAaManage.Show
    Unload Me
End If
End Sub

Private Sub txtPassword_Click()
    txtPassword.Text = ""
End Sub

Private Sub txtUsername_Click()
    txtUsername.Text = ""
End Sub
