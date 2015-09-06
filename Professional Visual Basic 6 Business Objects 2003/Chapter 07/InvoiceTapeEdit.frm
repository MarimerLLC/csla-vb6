VERSION 5.00
Begin VB.Form InvoiceTapeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rental tape"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPrice 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Price"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "InvoiceTapeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjInvoiceTape As InvoiceTape
Attribute mobjInvoiceTape.VB_VarHelpID = -1

Public Sub Component(InvoiceTapeObject As InvoiceTape)

  Set mobjInvoiceTape = InvoiceTapeObject

End Sub

Private Sub cmdApply_Click()

  mobjInvoiceTape.ApplyEdit
  mobjInvoiceTape.BeginEdit


End Sub

Private Sub cmdCancel_Click()

  mobjInvoiceTape.CancelEdit
  Unload Me


End Sub

Private Sub cmdOK_Click()

  mobjInvoiceTape.ApplyEdit
  Unload Me


End Sub

Private Sub Form_Load()

  mflgLoading = True
  With mobjInvoiceTape
     EnableOK .IsValid
    lblTitle = .Title
    txtPrice = .Price
    .BeginEdit
  End With
  mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid

End Sub

Private Sub mobjInvoiceTape_Valid(IsValid As Boolean)

  EnableOK IsValid

End Sub


