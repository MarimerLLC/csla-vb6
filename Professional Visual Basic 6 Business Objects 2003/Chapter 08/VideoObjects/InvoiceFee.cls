VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvoiceFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements InvoiceItem

Event Valid(IsValid As Boolean)

Private Type FeeProps
  InvoiceID As Long
  FeeID As Long
  VideoTitle As String * 30
  EnteredDate As Date
  DaysOver As Integer
  Paid As Boolean
  PaidDate As Date
End Type

Private mudtProps As FeeProps
Private mudtSave As FeeProps
Private mudtChild As FeeProps

'Private mobjParent As Tape

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChildEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit()

  If Not mflgChildEditing Then Err.Raise 445
  If mflgEditing Then Err.Raise 445
  
  ' save object state
  LSet mudtSave = mudtProps
  mflgEditing = True

End Sub

Public Sub CancelEdit()

  If Not mflgChildEditing Then Err.Raise 445
  If Not mflgEditing Then Err.Raise 445
  
  mflgEditing = False
  mflgDeleted = False
  mflgDirty = False
  ' restore object state
  LSet mudtProps = mudtSave
  
End Sub

Public Sub ApplyEdit()

  If Not mflgChildEditing Then Err.Raise 445
  If Not mflgEditing Then Err.Raise 445
  
  mflgEditing = False

End Sub

Private Sub Class_Initialize()

  mflgNew = True
  Set mobjValid = New BrokenRules
 ' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  '
  ' mobjValid.RuleBroken "RuleName", True
  ' mobjValid.RuleBroken "DateAcquired", True
  
End Sub

Public Property Get IsValid() As Boolean

  IsValid = (mobjValid.Count = 0)

End Property

Private Sub InvoiceItem_ApplyEdit()

  ApplyEdit
  
End Sub

Private Sub InvoiceItem_BeginEdit()

  BeginEdit
  
End Sub

Private Sub InvoiceItem_CancelEdit()

  CancelEdit
  
End Sub

Private Sub InvoiceItem_Delete()

  Delete
  
End Sub

Private Property Get InvoiceItem_IsDeleted() As Boolean

  InvoiceItem_IsDeleted = IsDeleted
  
End Property

Private Property Get InvoiceItem_IsDirty() As Boolean

  InvoiceItem_IsDirty = IsDirty
  
End Property

Private Property Get InvoiceItem_IsNew() As Boolean

  InvoiceItem_IsNew = IsNew
  
End Property

Private Property Get InvoiceItem_IsValid() As Boolean

  InvoiceItem_IsValid = IsValid
  
End Property

Private Property Get InvoiceItem_ItemDescription() As String

  InvoiceItem_ItemDescription = VideoTitle
  
End Property

Private Property Get InvoiceItem_ItemType() As ItemTypes

  InvoiceItem_ItemType = ITEM_FEE
  
End Property

Private Property Get InvoiceItem_Total() As Double

  InvoiceItem_Total = Fee
  
End Property

Private Sub mobjValid_BrokenRule()

  RaiseEvent Valid(False)

End Sub

Private Sub mobjValid_NoBrokenRules()

  RaiseEvent Valid(True)

End Sub

Public Sub Delete()

  If Not mflgChildEditing Then Err.Raise 445
  If Not mflgEditing Then Err.Raise 445
  
  mflgDeleted = True
  mflgDirty = True
  
End Sub

Public Property Get IsDeleted() As Boolean

  IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

  IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

  IsDirty = mflgDirty Or mflgNew

End Property

Friend Sub ChildBeginEdit()

  If mflgChildEditing Then Err.Raise 445
  
  ' save object state
  LSet mudtChild = mudtProps
  mflgChildEditing = True

End Sub

Friend Sub ChildCancelEdit()

  If Not mflgChildEditing Then Err.Raise 445
  
  mflgChildEditing = False
  mflgDeleted = False
  ' restore object state
  LSet mudtProps = mudtChild

End Sub

Friend Sub ChildApplyEdit(InvoiceID As Long)

  If Not mflgChildEditing Then Err.Raise 445
  
  If Not mflgDeleted And Not mflgNew Then _
    Pay
  If mflgDeleted And Not mflgNew Then
    ' code to delete the object's data goes here
    DeleteObject mudtProps.FeeID
    mflgNew = True
    mflgDeleted = False

  ElseIf mflgDirty Or mflgNew Then
    If Not IsValid Then Err.Raise 445
    mudtProps.InvoiceID = InvoiceID
    ' save object to database if appropriate
    Save
    'If mudtProps.LateFee Then
    '  ' add the late fee to the customer
    'End If
    ' save object state
    LSet mudtChild = mudtProps
    mflgNew = False

  End If

  mflgDirty = False
  mflgChildEditing = False

End Sub

Friend Sub ChildLoad(FeeID As Long)

  If mflgChildEditing Then Err.Raise 445
  If Not mflgNew Then Err.Raise 445
  
' code to load the object goes here
  Fetch FeeID
  mflgNew = False

End Sub

Public Property Get FeeID() As Long

  FeeID = mudtProps.FeeID

End Property

Public Property Get VideoTitle() As String

  VideoTitle = Trim$(mudtProps.VideoTitle)

End Property

Public Property Get DaysOver() As Integer

  DaysOver = mudtProps.DaysOver

End Property

Public Property Get Fee() As Double

  Fee = DaysOver * 1.5

End Property

Public Property Get EnteredDate() As Date

  EnteredDate = mudtProps.EnteredDate

End Property

Private Sub Pay()
  
  With mudtProps
    .Paid = True
    .PaidDate = Now
  End With
  
  mflgDirty = True

End Sub

Private Sub Fetch(FeeID As Long)

  Dim rsLateFee As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT LateFee.*, Video.Title " & _
    "FROM (LateFee INNER JOIN Tape ON " & _
    "LateFee.TapeID = Tape.TapeID) " & _
    "INNER JOIN Video ON Tape.VideoID = Video.VideoID " & _
    "WHERE LateFee.FeeID=" & Format$(FeeID)
  Set rsLateFee = New Recordset
  rsLateFee.Open strSQL, cnVideo

  With rsLateFee
    mudtProps.FeeID = .Fields("FeeID")
    mudtProps.VideoTitle = .Fields("Title")
    mudtProps.EnteredDate = .Fields("FeeDate")
    mudtProps.DaysOver = .Fields("DaysOver")
    mudtProps.Paid = .Fields("Paid")
    If Not IsNull(.Fields("PaidDate")) Then _
      mudtProps.PaidDate = .Fields("PaidDate")
    .Close
  End With

  Set rsLateFee = Nothing

End Sub

Private Sub DeleteObject(FeeID As Long)

  Dim rsLateFee As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT DELETED FROM LateFee WHERE FeeID=" & FeeID
  Set rsLateFee = New Recordset
  rsLateFee.Open strSQL, cnVideo

  With rsLateFee
    .Fields("Deleted") = True
    .Update
    .Close
  End With

End Sub

Private Sub Save()

  Dim rsLateFee As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT * FROM LateFee WHERE FeeID=" & _
    mudtProps.FeeID
  Set rsLateFee = New Recordset
  rsLateFee.Open strSQL, cnVideo, , adLockOptimistic
  If mflgNew Then rsLateFee.AddNew

  With rsLateFee
    .Fields("Paid") = mudtProps.Paid
    .Fields("PaidDate") = mudtProps.PaidDate
    .Update
    If mflgNew Then mudtProps.FeeID = .Fields("FeeID")
    .Close
  End With

  Set rsLateFee = Nothing

End Sub

