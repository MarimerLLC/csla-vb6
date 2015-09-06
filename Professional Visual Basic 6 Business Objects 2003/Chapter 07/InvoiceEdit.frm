VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form InvoiceEdit 
   Caption         =   "Invoice"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   9000
   Begin VB.Frame Frame1 
      Caption         =   "Invoice items"
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   8775
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   7560
         TabIndex        =   10
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   6360
         TabIndex        =   9
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   2640
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lblPhone 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblInvoiceID 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Customer Phone"
      Height          =   255
      Left            =   105
      TabIndex        =   5
      Top             =   1110
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Customer Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   630
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Invoice Number"
      Height          =   255
      Left            =   135
      TabIndex        =   3
      Top             =   150
      Width           =   1335
   End
End
Attribute VB_Name = "InvoiceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjInvoice As Invoice
Attribute mobjInvoice.VB_VarHelpID = -1

Public Sub Component(InvoiceObject As Invoice)

  Set mobjInvoice = InvoiceObject

End Sub

Private Sub cmdAdd_Click()
  Dim frmItem As InvoiceTapeEdit
  Dim strID As String
  
  strID = InputBox$("Scan the tape ID", "Tape ID")
  If Val(strID) = 0 Then Exit Sub
  Set frmItem = New InvoiceTapeEdit
  On Error GoTo ADDERR
  frmItem.Component mobjInvoice.InvoiceItems.Add(Val(strID))
  frmItem.Show vbModal
  LoadItems
  Exit Sub
ADDERR:
  If Err = vbObjectError + 1100 Then
    MsgBox "Tape is already checked out", vbExclamation

  Else
    MsgBox "Invalid tape id", vbExclamation

  End If

  Set frmItem = Nothing


End Sub

Private Sub cmdApply_Click()

  mobjInvoice.ApplyEdit
  mobjInvoice.BeginEdit


End Sub

Private Sub cmdCancel_Click()

  mobjInvoice.CancelEdit
  Unload Me


End Sub

Private Sub cmdEdit_Click()

  Dim frmItem As InvoiceTapeEdit
  Dim objItem As InvoiceItem
  
  Set objItem = _
    mobjInvoice.InvoiceItems(Val(lvwItems.SelectedItem.Key))
  If objItem.ItemType = ITEM_TAPE Then
    Set frmItem = New InvoiceTapeEdit
    frmItem.Component _
      mobjInvoice.InvoiceItems(Val(lvwItems.SelectedItem.Key))
    frmItem.Show vbModal
    LoadItems

  Else
    MsgBox "Only tape rental items can be edited", _
      vbInformation

  End If

End Sub


Private Sub cmdOK_Click()

  mobjInvoice.ApplyEdit
  Unload Me


End Sub

Private Sub cmdRemove_Click()
  mobjInvoice.InvoiceItems.Remove Val(lvwItems.SelectedItem.Key)
  LoadItems

End Sub

Private Sub Form_Load()

  mflgLoading = True
  With mobjInvoice
     EnableOK .IsValid
    If .IsNew Then
      lblInvoiceID = "(new)"

    Else
      lblInvoiceID = .InvoiceID

    End If

    lblName = .CustomerName
    lblPhone = .CustomerPhone
    .BeginEdit
  End With
  LoadItems
  mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid

End Sub

Private Sub mobjInvoice_Valid(IsValid As Boolean)

  EnableOK IsValid

End Sub

Private Sub LoadItems()

  Dim objItem As InvoiceItem
  Dim itmList As ListItem
  Dim lngIndex As Long
  
  lvwItems.ListItems.Clear

  For lngIndex = 1 To mobjInvoice.InvoiceItems.Count
    Set itmList = lvwItems.ListItems.Add _
      (Key:=Format$(lngIndex) & "K")
    Set objItem = mobjInvoice.InvoiceItems(lngIndex)

    With itmList
      .Text = IIf(objItem.ItemType = ITEM_FEE, _
        "Late fee", "Rental")
      If objItem.IsDeleted Then .Text = "(deleted)"
      .SubItems(1) = objItem.ItemDescription
      .SubItems(2) = Format$(objItem.Total, "0.00")
    End With

  Next

End Sub


