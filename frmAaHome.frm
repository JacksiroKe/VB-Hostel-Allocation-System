VERSION 5.00
Begin VB.Form frmAaHome 
   Caption         =   "Eldoret Polytechnic Accomodation MIS"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "News Gothic"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAaHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddNew_Click()
    frmAaManage.Show
    Unload Me
End Sub

Private Sub Command1_Click()
    frmAaEvaluate.Show
    Unload Me
End Sub


