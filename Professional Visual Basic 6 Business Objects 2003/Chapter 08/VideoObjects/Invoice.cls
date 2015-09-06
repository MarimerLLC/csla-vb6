VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private Type InvoiceProps
  InvoiceID As Long
  CustomerID As Long
  CustomerName As String * 50
  CustomerPhone As String * 25
End Type

Private mudtProps As InvoiceProps
Private mudtSave As InvoiceProps

Private mobjItems As InvoiceItems

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit()

  If mflgEditing Then Err.Raise 445
  
  ' save object state
  LSet mudtSave = mudtProps
  mflgEditing = True
  mobjItems.BeginEdit
  
End Sub

Public Sub CancelEdit()

  If Not mflgEditing Then Err.Raise 445
  
  mflgEditing = False
  mflgDeleted = False
  mflgDirty = False
  ' restore object state
  LSet mudtProps = mudtSave
  mobjItems.CancelEdit

End Sub

Public Sub ApplyEdit()

  If Not mflgEditing Then Err.Raise 445

  If mflgDeleted And Not mflgNew Then
    ' code to delete the object's data goes here
    DeleteObject mudtProps.InvoiceID
    mflgNew = True
    mflgDeleted = False

  ElseIf mflgDirty Or mflgNew Then
    If Not IsValid Then Err.Raise 445
    ' save object to database if appropriate
    Save
    ' save object state
    LSet mudtSave = mudtProps
    mflgNew = False

  End If

  mflgDirty = False
  mflgEditing = False
  mobjItems.ApplyEdit mudtProps.InvoiceID
  
End Sub

Private Sub Class_Initialize()

  mflgNew = True
  Set mobjValid = New BrokenRules
 
  Set mobjItems = New InvoiceItems
 ' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  '
  ' mobjValid.RuleBroken "RuleName", True
  mobjValid.RuleBroken "Customer", True
  
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

Public Sub Load(InvoiceID As Long)

  If mflgEditing Then Err.Raise 445
  If Not mflgNew Then Err.Raise 445
  
 ' code to load the object goes here
  Fetch InvoiceID
  mflgNew = False
  
  mobjValid.RuleBroken "Customer", False
  mobjItems.Load mudtProps.InvoiceID

End Sub

Public Sub Delete()

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

Public Property Get InvoiceID() As Long

  InvoiceID = mudtProps.InvoiceID

End Property

Public Property Get InvoiceItems() As InvoiceItems

  Set InvoiceItems = mobjItems

End Property

Public Property Get CustomerID() As Long

  CustomerID = mudtProps.CustomerID

End Property

Public Property Get CustomerName() As String

  CustomerName = Trim$(mudtProps.CustomerName)

End Property

Public Property Get CustomerPhone() As String

  CustomerPhone = Trim$(mudtProps.CustomerPhone)

End Property

Public Property Get SubTotal() As Double

  SubTotal = mobjItems.SubTotal

End Property

Public Property Get Tax() As Double

  Tax = Val(Format$(mobjItems.SubTotal * 0.065 + 0.0005, "0.00"))

End Property

Public Property Get Total() As Double

  Total = SubTotal + Tax

End Property

Friend Sub Initialize(Renter As Customer)

  With Renter
    mudtProps.CustomerID = .CustomerID
    mobjValid.RuleBroken "Customer", .IsNew
    mudtProps.CustomerName = .Name
    mudtProps.CustomerPhone = .Phone
    mobjItems.LoadFees .CustomerID
  End With

End Sub

Private Sub Save()

  Dim rsInvoice As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT * FROM Invoice WHERE InvoiceID=" & InvoiceID
  Set rsInvoice = New Recordset
  rsInvoice.Open strSQL, cnVideo, , adLockOptimistic
  If mflgNew Then rsInvoice.AddNew

  With rsInvoice
    .Fields("CustomerID") = mudtProps.CustomerID
    .Fields("Subtotal") = Me.SubTotal
    .Fields("Tax") = Me.Tax
    .Fields("Total") = Me.Total
    .Update
    If IsNew Then mudtProps.InvoiceID = .Fields("InvoiceID")
    .Close
  End With

  strSQL = "SELECT NAME,PHONE FROM Customer " & _
    "WHERE CustomerID=" & mudtProps.CustomerID
  rsInvoice.Open strSQL, cnVideo
  mudtProps.CustomerName = rsInvoice("Name")
  mudtProps.CustomerPhone = rsInvoice("Phone")
  rsInvoice.Close
  Set rsInvoice = Nothing

End Sub

Private Sub DeleteObject(InvoiceID As Long)

  cnVideo.Execute "DELETE FROM Invoice WHERE InvoiceID=" & _
    InvoiceID

End Sub

Private Sub Fetch(InvoiceID As Long)

  Dim rsInvoice As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Invoice.*,Customer.Name,Customer.Phone " & _
    "FROM Invoice INNER JOIN Customer ON " & _
    "Invoice.CustomerID = Customer.CustomerID " & _
    "WHERE InvoiceID=" & CStr(InvoiceID)
  Set rsInvoice = New Recordset
  rsInvoice.Open strSQL, cnVideo

  With rsInvoice
    mudtProps.InvoiceID = .Fields("InvoiceID")
    mudtProps.CustomerID = .Fields("CustomerID")
    mudtProps.CustomerName = .Fields("Name")
    mudtProps.CustomerPhone = .Fields("Phone")
    .Close
  End With

  Set rsInvoice = Nothing
  mobjValid.RuleBroken "Customer", False

End Sub

