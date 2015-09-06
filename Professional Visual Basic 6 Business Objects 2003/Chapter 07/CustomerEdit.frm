VERSION 5.00
Begin VB.Form CustomerEdit 
   Caption         =   "Customer"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2505
   ScaleWidth      =   4185
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtZipCode 
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   2640
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtAddr2 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtAddr1 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Phone"
      Height          =   255
      Left            =   105
      TabIndex        =   8
      Top             =   1590
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Address"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "CustomerEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjCustomer As Customer
Attribute mobjCustomer.VB_VarHelpID = -1

Public Sub Component(CustomerObject As Customer)

  Set mobjCustomer = CustomerObject

End Sub

Private Sub cmdApply_Click()

  mobjCustomer.ApplyEdit
  mobjCustomer.BeginEdit


End Sub

Private Sub cmdCancel_Click()

  mobjCustomer.CancelEdit
  Unload Me


End Sub

Private Sub cmdOK_Click()

  mobjCustomer.ApplyEdit
  Unload Me


End Sub

Private Sub Form_Load()

  mflgLoading = True
  With mobjCustomer
    EnableOK .IsValid
    If .IsNew Then
      Caption = "Customer [(new)]"

    Else
      Caption = "Customer [" & .Name & "]"

    End If

    txtName = .Name
    txtAddr1 = .Address1
    txtAddr2 = .Address2
    txtCity = .City
    txtState = .State
    txtZipCode = .ZipCode
    txtPhone = .Phone
    .BeginEdit
  End With
  
  mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid

End Sub

Private Sub mobjCustomer_Valid(IsValid As Boolean)

  EnableOK IsValid

End Sub

Private Sub txtName_Change()

  If Not mflgLoading Then _
    TextChange txtName, mobjCustomer, "Name"

End Sub

Private Sub txtName_LostFocus()

  txtName = TextLostFocus(txtName, mobjCustomer, "Name")

End Sub

Private Sub txtAddr1_Change()

  If Not mflgLoading Then _
    TextChange txtAddr1, mobjCustomer, "Address1"

End Sub

Private Sub txtAddr1_LostFocus()

  txtAddr1 = TextLostFocus(txtAddr1, mobjCustomer, "Address1")

End Sub

Private Sub txtAddr2_Change()

  If Not mflgLoading Then _
    TextChange txtAddr2, mobjCustomer, "Address2"

End Sub

Private Sub txtAddr2_LostFocus()

  TtxtAddr2 = TextLostFocus(txtAddr2, mobjCustomer, "Address2")

End Sub

Private Sub txtCity_Change()

  If Not mflgLoading Then _
    TextChange txtCity, mobjCustomer, "City"

End Sub

Private Sub txtCity_LostFocus()

  txtCity = TextLostFocus(txtCity, mobjCustomer, "City")

End Sub

Private Sub txtPhone_Change()

  If Not mflgLoading Then _
    TextChange txtPhone, mobjCustomer, "Phone"

End Sub

Private Sub txtPhone_LostFocus()

  txtPhone = TextLostFocus(txtPhone, mobjCustomer, "Phone")

End Sub

Private Sub txtState_Change()

  If Not mflgLoading Then _
    TextChange txtState, mobjCustomer, "State"

End Sub

Private Sub txtState_LostFocus()

  txtState = TextLostFocus(txtState, mobjCustomer, "State")

End Sub

Private Sub txtZipCode_Change()

  If Not mflgLoading Then _
    TextChange txtZipCode, mobjCustomer, "ZipCode"

End Sub

Private Sub txtZipCode_LostFocus()

  txtZipCode = TextLostFocus(txtZipCode, mobjCustomer, "ZipCode")

End Sub


