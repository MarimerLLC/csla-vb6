VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvoiceItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolItems As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

  Set mcolItems = New Collection
  
End Sub

Public Function Item(ByVal Index As Variant) As InvoiceItem
Attribute Item.VB_UserMemId = 0

  Set Item = mcolItems(Index)

End Function

Public Function Count() As Long

  Count = mcolItems.Count
  
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

  Set NewEnum = mcolItems.[_NewEnum]
  
End Function

Public Sub BeginEdit()

  Dim objInvoiceItem As InvoiceItem
  Dim objTape As InvoiceTape
  Dim objFee As InvoiceFee
  
  If mflgEditing Then Err.Raise 445
  
  For Each objInvoiceItem In mcolItems
    If objInvoiceItem.ItemType = ITEM_FEE Then
      Set objFee = objInvoiceItem
      objFee.ChildBeginEdit

    Else
      Set objTape = objInvoiceItem
      objTape.ChildBeginEdit

    End If
  
  Next
  
  mflgEditing = True

End Sub

Public Sub CancelEdit()

  Dim objInvoiceItem As InvoiceItem
  Dim objTape As InvoiceTape
  Dim objFee As InvoiceFee
  Dim lngIndex As Long
  
  If Not mflgEditing Then Err.Raise 445
  
  For lngIndex = mcolItems.Count To 1 Step -1
    Set objInvoiceItem = mcolItems.Item(lngIndex)
    If objInvoiceItem.ItemType = ITEM_FEE Then
      Set objFee = objInvoiceItem

      With objFee
        If .IsNew Then _
          mcolItems.Remove lngIndex
        .ChildCancelEdit
      End With

    Else
      Set objTape = objInvoiceItem

      With objTape
        .ChildCancelEdit
        If .IsNew Then _
          mcolItems.Remove lngIndex
      End With

    End If

  Next lngIndex

End Sub

Public Sub ApplyEdit(InvoiceID As Long)

  Dim objInvoiceItem As InvoiceItem
  Dim objTape As InvoiceTape
  Dim objFee As InvoiceFee
  Dim lngIndex As Long

  If Not mflgEditing Then Err.Raise 445

  For lngIndex = mcolItems.Count To 1 Step -1
    Set objInvoiceItem = mcolItems.Item(lngIndex)
    If objInvoiceItem.ItemType = ITEM_FEE Then
      Set objFee = objInvoiceItem

      With objFee
        If .IsDeleted Then _
          mcolItems.Remove lngIndex
        .ChildApplyEdit InvoiceID
      End With

    Else
      Set objTape = objInvoiceItem

      With objTape
        If .IsDeleted Then _
          mcolItems.Remove lngIndex
        .ChildApplyEdit InvoiceID
      End With

    End If

  Next lngIndex

  mflgEditing = False

End Sub

Public Sub Delete()

  Dim objInvoiceItem As InvoiceItem
  
  If Not mflgEditing Then Err.Raise 445
  ' code to delete the object's data goes here

  For Each objInvoiceItem In mcolItems
    With objInvoiceItem
      .BeginEdit
      .Delete
      .ApplyEdit
    End With

  Next

End Sub

Public Function Add(TapeID As Long) As InvoiceItem

  Dim objInvoiceTape As InvoiceTape
  Dim objTape As Tape
  
  If Not mflgEditing Then Err.Raise 445

  Set objTape = New Tape
  objTape.Load TapeID
  Set objInvoiceTape = New InvoiceTape
  objInvoiceTape.Initialize objTape, Me
  If mflgEditing Then objInvoiceTape.ChildBeginEdit
  
  Set Add = objInvoiceTape
  Set objInvoiceTape = Nothing

End Function

Public Sub Remove(ByVal Index As Variant)

  If Not mflgEditing Then Err.Raise 445

  With mcolItems(Index)
    .BeginEdit
    .Delete
    .ApplyEdit
  End With

End Sub

Friend Sub Load(InvoiceID As Long)

  If mflgEditing Then Err.Raise 445
  
  ' here we'll put looping code to load
  ' each of the InvoiceItem objects for the InvoiceID
  
  ' For Each InvoiceItemID
  '   Set objInvoiceItem = New InvoiceItem
  '   objInvoiceItem.ChildLoad InvoiceItemID
  '   mcolItems.Add Item:=objInvoiceItem

  ' Loop

End Sub

Friend Sub AddTape(Child As InvoiceTape)

  mcolItems.Add Item:=Child

End Sub

Friend Sub LoadFees(CustomerID As Long)

  If mflgEditing Then Err.Raise 445

  ' here we'll put looping code to create
  ' each of the InvoiceFee objects for the CustomerID
  
  ' For Each InvoiceFeeID
  '   Set objInvoiceFee = New InvoiceFee
  '   objInvoiceFee.ChildLoad InvoiceFeeID
  '   mcolItems.Add Item:=objInvoiceFee
  ' Loop

End Sub

Friend Property Get SubTotal() As Double

  Dim objItem As InvoiceItem
  Dim dblAmount As Double
  
  For Each objItem In mcolItems
    If Not objItem.IsDeleted Then _
      dblAmount = dblAmount + objItem.Total
  Next

  SubTotal = dblAmount

End Property


