VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private Type TapeProps
  TapeID As Long
  VideoID As Long
  Title As String * 30
  CheckedOut As Boolean
  DateAcquired As Variant
  DateDue As Variant
  LateFee As Boolean
  InvoiceID As Long
End Type

Private mudtProps As TapeProps
Private mudtSave As TapeProps
Private mudtChild As TapeProps
Private mobjParent As Tapes
Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChild As Boolean
Private mflgChildEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit()

  If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
  If mflgEditing Then Err.Raise 445
  
  ' save object state
  LSet mudtSave = mudtProps
  mflgEditing = True

End Sub

Public Sub CancelEdit()

  If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
  If Not mflgEditing Then Err.Raise 445
  
  mflgEditing = False
  mflgDeleted = False
  mflgDirty = False
  ' restore object state
  LSet mudtProps = mudtSave
  
End Sub

Public Sub ApplyEdit()

  If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
  If Not mflgChild And mflgNew Then Err.Raise 445
  If Not mflgEditing Then Err.Raise 445

  If Not mflgChild Then
    If mflgDeleted And Not mflgNew Then
      ' code to delete the object's data goes here
      DeleteObject mudtProps.TapeID
      mflgNew = True
      mflgDeleted = False

    ElseIf mflgDirty Or mflgNew Then
      If Not IsValid Then Err.Raise 445
      ' save object to database if appropriate
      Save
      If mudtProps.LateFee Then
        ' add the late fee to the customer
        SaveLateFee
      End If
      
      ' save object state
      LSet mudtSave = mudtProps
      mflgNew = False

    End If
    
    mflgDirty = False

  Else
    If Not mobjParent Is Nothing Then
      mobjParent.AddTape Me
      Set mobjParent = Nothing
    End If

  
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
  mobjValid.RuleBroken "DateAcquired", True
  
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

Public Sub Load(TapeID As Long)

  If mflgChild Then Err.Raise 445
  If mflgEditing Then Err.Raise 445
  If Not mflgNew Then Err.Raise 445
  
 ' code to load the object goes here
  Fetch TapeID
  mflgNew = False

End Sub

Public Sub Delete()

  If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
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

Friend Sub SetAsChild(objTapes As Tapes)

  If mflgNew Then Set mobjParent = objTapes
  mflgChild = True

End Sub

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

Friend Sub ChildApplyEdit(VideoID As Long)

  If Not mflgChildEditing Then Err.Raise 445

  If mflgDeleted And Not mflgNew Then
    ' code to delete the object's data goes here
    DeleteObject mudtProps.TapeID
    mflgNew = True
    mflgDeleted = False

  ElseIf mflgDirty Or mflgNew Then
    If Not IsValid Then Err.Raise 445
    mudtProps.VideoID = VideoID
    ' save object to database if appropriate
    Save
    If mudtProps.LateFee Then
      ' add the late fee to the customer
      SaveLateFee
    End If
    ' save object state
    LSet mudtChild = mudtProps
    mflgNew = False

  End If

  mflgDirty = False
  mflgChildEditing = False

End Sub

Friend Sub ChildLoad(TapeID As Long)

  If mflgChildEditing Then Err.Raise 445
  If Not mflgNew Then Err.Raise 445
  
' code to load the object goes here
  Fetch TapeID
  mflgNew = False

End Sub

Public Property Get TapeID() As Long

  TapeID = mudtProps.TapeID

End Property

Public Property Get Title() As String

  If mflgNew Then Err.Raise 445
  
  Title = Trim$(mudtProps.Title)

End Property

Public Property Get CheckedOut() As Boolean

  CheckedOut = mudtProps.CheckedOut

End Property

Public Property Let DateAcquired(Value As Variant)

  If Not mflgEditing Then Err.Raise 383
  If Not mflgNew Then Err.Raise 383
  
  If Len(Trim(Value)) = 0 Then
    mudtProps.DateAcquired = ""
    mobjValid.RuleBroken "DateAcquired", True

  ElseIf IsDate(Value) Then
    mudtProps.DateAcquired = CVDate(Value)
    mobjValid.RuleBroken "DateAcquired", False

  Else
    mobjValid.RuleBroken "DateAcquired", True

  End If

  mflgDirty = True

End Property

Public Property Get DateAcquired() As Variant

  DateAcquired = mudtProps.DateAcquired

End Property

Public Property Get Invoice() As Invoice

  Dim objInvoice As Invoice
  
  If Not mudtProps.CheckedOut Then Err.Raise 445
  
  Set objInvoice = New Invoice
  objInvoice.Load mudtProps.InvoiceID
  Set Invoice = objInvoice
  Set Invoice = Nothing

End Property

Friend Sub CheckOut(InvoiceID As Long)

  If Not mflgEditing Then Err.Raise 445

  With mudtProps
    .CheckedOut = True
    .DateDue = DateAdd("d", 2, Now)
    .InvoiceID = InvoiceID
  End With

  mflgDirty = True

End Sub

Public Sub CheckIn()

  If Not mflgEditing Then Err.Raise 445
  
  With mudtProps
    .CheckedOut = False
    If DateDiff("d", .DateDue, Now) > 0 Then _
      .LateFee = True
 
  End With

  mflgDirty = True

End Sub

Public Property Get DateDue() As Variant

  If Not mudtProps.CheckedOut Then Err.Raise 438
  
  DateDue = mudtProps.DateDue

End Property

Private Sub Save()

  Dim rsTape As Recordset
  Dim rsVideo As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT * FROM Tape WHERE TapeID=" & TapeID
  Set rsTape = New Recordset
  rsTape.Open strSQL, cnVideo, , adLockOptimistic
  If mflgNew Then rsTape.AddNew

  With rsTape
    .Fields("VideoID") = mudtProps.VideoID
    .Fields("CheckedOut") = mudtProps.CheckedOut
    .Fields("DateAcquired") = mudtProps.DateAcquired
    .Fields("DateDue") = mudtProps.DateDue
    .Fields("InvoiceID") = mudtProps.InvoiceID
    .Update
    If mflgNew Then mudtProps.TapeID = .Fields("TapeID")
    .Close
  End With

  Set rsTape = Nothing
  If mflgNew Then
    strSQL = "SELECT TITLE FROM Video WHERE VideoID=" & _
      mudtProps.VideoID
    Set rsVideo = New Recordset
    rsVideo.Open strSQL, cnVideo
    mudtProps.Title = rsVideo("Title")
    rsVideo.Close
    Set rsVideo = Nothing
  End If

End Sub

Private Sub DeleteObject(TapeID As Long)

  cnVideo.Execute "DELETE FROM Tape WHERE TapeID=" & TapeID

End Sub

Private Sub Fetch(TapeID As Long)

  Dim rsTape As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Tape.*,Video.Title " & _
    "FROM Tape INNER JOIN Video ON " & _
    "Tape.VideoID = Video.VideoID " & _
    "WHERE TapeID=" & TapeID
  Set rsTape = New Recordset
  rsTape.Open strSQL, cnVideo

  With rsTape
    mudtProps.TapeID = .Fields("TapeID")
    mudtProps.VideoID = .Fields("VideoID")
    mudtProps.Title = .Fields("Title")
    mudtProps.CheckedOut = .Fields("CheckedOut")
    mudtProps.DateAcquired = .Fields("DateAcquired")
    mudtProps.DateDue = .Fields("DateDue")
    mudtProps.InvoiceID = .Fields("InvoiceID")

    mobjValid.RuleBroken "DateAcquired", False
    
    rsTape.Close
  End With

  Set rsTape = Nothing

End Sub

Private Sub SaveLateFee()

  Dim rsLateFee As Recordset
  Dim rsInvoice As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT CUSTOMERID FROM INVOICE WHERE INVOICEID=" & _
    mudtProps.InvoiceID
  Set rsInvoice = New Recordset
  rsInvoice.Open strSQL, cnVideo
  strSQL = "LateFee"
  Set rsLateFee = New Recordset
  rsLateFee.Open strSQL, cnVideo, , adLockOptimistic
  rsLateFee.AddNew

  With rsLateFee
    .Fields("CustomerID") = rsInvoice("CustomerID")
    .Fields("TapeID") = mudtProps.TapeID
    .Fields("FeeDate") = Now
    .Fields("DaysOver") = DateDiff("d", mudtProps.DateDue, Now)
    .Fields("Paid") = False
    .Update
    .Close
  End With

  rsInvoice.Close
  Set rsInvoice = Nothing
  Set rsLateFee = Nothing

End Sub


