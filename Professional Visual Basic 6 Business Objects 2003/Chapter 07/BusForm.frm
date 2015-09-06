VERSION 5.00
Begin VB.Form BusForm 
   Caption         =   "BusForm"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "BusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private flgLoading As Boolean

Private WithEvents mobjBusiness As Business
Attribute mobjBusiness.VB_VarHelpID = -1

Public Sub Component(BusinessObject As Business)

  Set mobjBusiness = BusinessObject

End Sub

Private Sub cmdApply_Click()

  mobjBusiness.ApplyEdit
  mobjBusiness.BeginEdit


End Sub

Private Sub cmdCancel_Click()

  mobjBusiness.CancelEdit
  Unload Me


End Sub

Private Sub cmdOK_Click()

  mobjBusiness.ApplyEdit
  Unload Me


End Sub

Private Sub Form_Load()

  flgLoading = True
  With mobjBusiness
     EnableOK .IsValid
    ' load object values into form controls
    ' txtText = .Property
    .BeginEdit
  End With
  flgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid

End Sub

Private Sub mobjBusiness_Valid(IsValid As Boolean)

  EnableOK IsValid

End Sub

'Private Sub Text1_Change()
'
'  If flgLoading Then Exit Sub
'
'  mobjBusiness.Text1 = Text1.Text
'
'End Sub

'Private Sub Text1_LostFocus()
'
'  Text1.Text = mobjBusiness.Text1
'
'End Sub

