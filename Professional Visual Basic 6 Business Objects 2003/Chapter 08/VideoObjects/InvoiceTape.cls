VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvoiceTape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements InvoiceItem

Event Valid(IsValid As Boolean)

Private Type InvoiceTapeProps
  InvoiceID As Long
  ItemID As Long
  TapeID As Long
  Title As String * 30
  Price As Double
End Type

Private mudtProps As InvoiceTapeProps
Private mudtSave As InvoiceTapeProps
Private mudtChild As InvoiceTapeProps

Private mobjParent As InvoiceItems

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
  
  If Not mobjParent Is Nothing Then
    mobjParent.AddTape Me
    Set mobjParent = Nothing
  End If

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

  If mflgDeleted And Not mflgNew Then
    ' code to delete the object's data goes here
    DeleteObject mudtProps.ItemID
    mflgNew = True
    mflgDeleted = False

  ElseIf mflgDirty Or mflgNew Then
    If Not IsValid Then Err.Raise 445
    mudtProps.InvoiceID = InvoiceID
    ' save object to database if appropriate
    Save InvoiceID
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

Friend Sub ChildLoad(ItemID As Long)

  If mflgChildEditing Then Err.Raise 445
  If Not mflgNew Then Err.Raise 445
  
' code to load the object goes here
  Fetch ItemID
  mflgNew = False

End Sub

Public Property Get ItemID() As Long

  ItemID = mudtProps.ItemID

End Property

Public Property Get Title() As String

  Title = Trim$(mudtProps.Title)

End Property

Public Property Let Price(Value As Double)

  mudtProps.Price = Value
  mobjValid.RuleBroken "Price", (Value < 0 Or Value > 4)
  mflgDirty = True

End Property

Public Property Get Price() As Double

  Price = mudtProps.Price

End Property

Friend Sub Initialize(RentalTape As Tape, Parent As InvoiceItems)

  If mflgNew Then Set mobjParent = Parent

  With RentalTape
    mudtProps.TapeID = .TapeID
    mudtProps.Title = .Title
    mudtProps.Price = 1.99
  End With

End Sub

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

  InvoiceItem_ItemDescription = Title

End Property

Private Property Get InvoiceItem_Total() As Double

  InvoiceItem_Total = Price

End Property

Private Property Get InvoiceItem_ItemType() As ItemTypes

  InvoiceItem_ItemType = ITEM_TAPE

End Property

Private Sub Save(InvoiceID As Long)

  Dim rsInvoiceTape As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT * FROM InvoiceTape WHERE ItemID=" & _
    Format$(ItemID)
  Set rsInvoiceTape = New Recordset
  rsInvoiceTape.Open strSQL, cnVideo, , adLockOptimistic
  If mflgNew Then rsInvoiceTape.AddNew

  With rsInvoiceTape
    .Fields("InvoiceID") = InvoiceID
    .Fields("TapeID") = mudtProps.TapeID
    .Fields("Price") = mudtProps.Price
    .Update
    If IsNew Then mudtProps.ItemID = .Fields("ItemID")
    .Close
  End With

  strSQL = "SELECT Video.Title " & _
    "FROM Tape INNER JOIN Video ON " & _
    "Tape.VideoID = Video.VideoID " & _
    "WHERE Tape.TapeID=" & mudtProps.TapeID
  rsInvoiceTape.Open strSQL, cnVideo
  mudtProps.Title = rsInvoiceTape("Title")
  rsInvoiceTape.Close
  Set rsInvoiceTape = Nothing
  CheckOut InvoiceID
  
End Sub

Private Sub CheckOut(InvoiceID As Long)

  Dim objTape As Tape
  
  Set objTape = New Tape
  objTape.Load mudtProps.TapeID

  With objTape
    .BeginEdit
    .CheckOut InvoiceID
    .ApplyEdit
  End With

  Set objTape = Nothing

End Sub

Private Sub DeleteObject(ItemID As Long)

  cnVideo.Execute "DELETE FROM InvoiceTape WHERE ItemID=" & _
    CStr(ItemID)

End Sub

Private Sub Fetch(ItemID As Long)

  Dim rsInvoiceTape As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT InvoiceTape.*, Video.Title " & _
    "FROM (InvoiceTape INNER JOIN Tape ON " & _
    "InvoiceTape.TapeID = Tape.TapeID) " & _
    "INNER JOIN Video ON Tape.VideoID = Video.VideoID " & _
    "WHERE InvoiceTape.ItemID=" & ItemID
  Set rsInvoiceTape = New Recordset
  rsInvoiceTape.Open strSQL, cnVideo

  With rsInvoiceTape
    mudtProps.ItemID = .Fields("ItemID")
    mudtProps.TapeID = .Fields("TapeID")
    mudtProps.Title = .Fields("Title")
    mudtProps.Price = .Fields("Price")
    .Close
  End With

  Set rsInvoiceTape = Nothing

End Sub

