VERSION 5.00
Begin VB.Form frmEditPerson 
   Caption         =   "Edit Person"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3360
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2160
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   390
      Left            =   960
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtBirthDate 
      Height          =   330
      Left            =   960
      TabIndex        =   4
      Top             =   1185
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   960
      TabIndex        =   2
      Top             =   615
      Width           =   3375
   End
   Begin VB.TextBox txtSSN 
      Height          =   330
      Left            =   960
      TabIndex        =   0
      Top             =   105
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Age"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1755
      Width           =   735
   End
   Begin VB.Label lblAge 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   960
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Birthdate"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1245
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   690
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "SSN"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "frmEditPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjPerson As Person
Attribute mobjPerson.VB_VarHelpID = -1
Private mflgLoading As Boolean

Private Sub cmdApply_Click()

  mobjPerson.ApplyEdit
  mobjPerson.BeginEdit

End Sub

Private Sub cmdCancel_Click()

  ' do not save the object
  mobjPerson.CancelEdit
  Unload Me
 
End Sub

Private Sub cmdOK_Click()

  mobjPerson.ApplyEdit
  Unload Me

End Sub

Private Sub Form_Load()

  Dim strSSN As String

  Set mobjPerson = New Person
  strSSN = InputBox$("Enter the SSN")

  mobjPerson.Load (strSSN)

  mflgLoading = True
 
  With mobjPerson
    txtSSN = .SSN
    txtName = .Name
    txtBirthDate = .BirthDate
    lblAge = .Age
  End With
  mflgLoading = False

  EnableOK mobjPerson.IsValid
  mobjPerson.BeginEdit
  
End Sub

Private Sub mobjPerson_NewAge()

  lblAge = mobjPerson.Age

End Sub

Private Sub txtBirthdate_Change()

  If mflgLoading Then Exit Sub
  If IsDate(txtBirthDate) Then mobjPerson.BirthDate = txtBirthDate

End Sub

Private Sub txtName_Change()

  If mflgLoading Then Exit Sub
  mobjPerson.Name = txtName

End Sub

Private Sub txtSSN_Change()

  If mflgLoading Then Exit Sub
  mobjPerson.SSN = txtSSN
  
End Sub

Private Sub EnableOK(IsOK As Boolean)

  cmdOK.Enabled = IsOK
  cmdApply.Enabled = IsOK

End Sub

Private Sub mobjPerson_Valid(IsValid As Boolean)

  EnableOK IsValid

End Sub

