VERSION 5.00
Begin VB.Form TapeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tape"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAcquired 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblRented 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Rented out"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1635
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Purchase date"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Tape ID"
      Height          =   255
      Left            =   135
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblTapeID 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "TapeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjTape As Tape
Attribute mobjTape.VB_VarHelpID = -1

Public Sub Component(TapeObject As Tape)

  Set mobjTape = TapeObject

End Sub

Private Sub cmdApply_Click()

  mobjTape.ApplyEdit
  mobjTape.BeginEdit


End Sub

Private Sub cmdCancel_Click()

  mobjTape.CancelEdit
  Unload Me


End Sub

Private Sub cmdOK_Click()

  mobjTape.ApplyEdit
  Unload Me


End Sub

Private Sub Form_Load()

  mflgLoading = True
  With mobjTape
     EnableOK .IsValid
    If .IsNew Then
      Caption = "Tape [(new)]"

    Else
      Caption = "Tape [" & .Title & "]"
      txtAcquired.Locked = True
      txtAcquired.BackColor = lblTapeID.BackColor

    End If

    lblTapeID = .TapeID
    txtAcquired = .DateAcquired
    lblRented = IIf(.CheckedOut, "Yes", "No")
    .BeginEdit
  End With
  mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid

End Sub

Private Sub mobjTape_Valid(IsValid As Boolean)

  EnableOK IsValid

End Sub

Private Sub txtAcquired_Change()

  If Not mflgLoading Then _
    TextChange txtAcquired, mobjTape, "DateAcquired"

End Sub

Private Sub txtAcquired_LostFocus()

  TextLostFocus txtAcquired, mobjTape, "DateAcquired"

End Sub


