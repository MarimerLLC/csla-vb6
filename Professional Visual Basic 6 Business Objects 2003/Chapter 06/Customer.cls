VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private Type CustomerProps
  CustomerID As Long
  Name As String * 50
  Phone As String * 25
  Address1 As String * 30
  Address2 As String * 30
  City As String * 20
  State As String * 2
  ZipCode As String * 10
End Type

Private mudtProps As CustomerProps
Private mudtSave As CustomerProps

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

End Sub

Public Sub CancelEdit()

  If Not mflgEditing Then Err.Raise 445
  
  mflgEditing = False
  mflgDeleted = False
  mflgDirty = False
  ' restore object state
  LSet mudtProps = mudtSave
  
End Sub

Public Sub ApplyEdit()

  If Not mflgEditing Then Err.Raise 445

  If mflgDeleted And Not mflgNew Then
    ' code to delete the object's data goes here
    mflgNew = True
    mflgDeleted = False

  ElseIf mflgDirty Or mflgNew Then
    If Not IsValid Then Err.Raise 445
    ' save object to database if appropriate
    ' save object state
    LSet mudtSave = mudtProps
    mflgNew = False

  End If

  mflgDirty = False
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
  mobjValid.RuleBroken "Name", True
  mobjValid.RuleBroken "Phone", True

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

Public Sub Load(CustomerID As Long)

  If mflgEditing Then Err.Raise 445
  If Not mflgNew Then Err.Raise 445
  
 ' code to load the object goes here

  mflgNew = False

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

Public Property Get CustomerID() As Long

  CustomerID = mudtProps.CustomerID

End Property

Public Property Let Name(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtProps.Name) Then _
    Err.Raise vbObjectError + 1001, "String value too long"
  mobjValid.RuleBroken "Name", (Len(Trim$(Value)) = 0)

  mudtProps.Name = Value
  mflgDirty = True

End Property

Public Property Get Name() As String

  Name = Trim$(mudtProps.Name)

End Property

Public Property Let Address1(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtProps.Address1) Then _
    Err.Raise vbObjectError + 1001, "String value too long"

  mudtProps.Address1 = Value
  mflgDirty = True

End Property

Public Property Get Address1() As String

  Address1 = Trim$(mudtProps.Address1)

End Property

Public Property Let Address2(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtProps.Address2) Then _
    Err.Raise vbObjectError + 1001, "String value too long"
  
  mudtProps.Address2 = Value
  mflgDirty = True

End Property

Public Property Get Address2() As String

  Address2 = Trim$(mudtProps.Address2)

End Property

Public Property Let City(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtProps.City) Then _
    Err.Raise vbObjectError + 1001, "String value too long"
  
  mudtProps.City = Value
  mflgDirty = True

End Property

Public Property Get City() As String

  City = Trim$(mudtProps.City)

End Property

Public Property Let State(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtProps.State) Then _
    Err.Raise vbObjectError + 1001, "String value too long"
  
  mudtProps.State = UCase$(Value)
  mflgDirty = True

End Property

Public Property Get State() As String

  State = Trim$(mudtProps.State)

End Property

Public Property Let ZipCode(Value As String)

  Dim intDash As Integer

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtProps.ZipCode) Then _
    Err.Raise vbObjectError + 1001, "String value too long"
  
  intDash = InStr(1, Value, "-")

  If intDash > 0 Then
    mudtProps.ZipCode = _
      Format$(Val(Mid$(Value, 1, intDash - 1)), "00000") & _
      "-" & Format$(Val(Mid$(Value, intDash + 1)), "0000")

  Else
    mudtProps.ZipCode = _
      Format$(Val(Mid$(Value, 1, 5)), "00000") & _
      "-" & Format$(Val(Mid$(Value, 6)), "0000")

  End If

  If Right$(mudtProps.ZipCode, 5) = "-0000" Then _
    mudtProps.ZipCode = Left$(mudtProps.ZipCode, 5)

  mflgDirty = True

End Property

Public Property Get ZipCode() As String

  ZipCode = Trim$(mudtProps.ZipCode)

End Property

Public Property Let Phone(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtProps.Phone) Then _
    Err.Raise vbObjectError + 1001, "String value too long"

  mobjValid.RuleBroken "Phone", (Len(Trim$(Value)) = 0)
  
  mudtProps.Phone = Value
  mflgDirty = True

End Property

Public Property Get Phone() As String

  Phone = Trim$(mudtProps.Phone)

End Property

Public Function CreateInvoice() As Invoice

  Dim objInvoice As Invoice

  Set objInvoice = New Invoice
  objInvoice.Initialize Me
  Set CreateInvoice = objInvoice
  Set objInvoice = Nothing

End Function



