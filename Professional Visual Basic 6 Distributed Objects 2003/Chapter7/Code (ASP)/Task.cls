VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private mudtProps As TaskProps
Private mcolStack As Collection
Private mobjParent As Tasks

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit()
  ' save object state
  mcolStack.Add GetState
End Sub

Public Sub CancelEdit()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = False
  ' restore object state
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
End Sub

Public Sub ApplyEdit()
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddTask Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
End Sub

Friend Sub Initialize(Parent As Tasks)
  If mudtProps.IsNew Then Set mobjParent = Parent
End Sub

Private Sub Class_Initialize()
  Set mcolStack = New Collection
  Set mobjValid = New BrokenRules
  mudtProps.IsNew = True
  ' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  '
  ' mobjValid.RuleBroken "RuleName", True
End Sub

Private Sub Class_Terminate()
  If mobjParent Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
    
  Set mcolStack = Nothing
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

Public Sub Load()
  If mcolStack.Count > 0 Then Err.Raise 445
  If Not mudtProps.IsNew Then Err.Raise 445
  
  mudtProps.IsNew = False

  ' code to load the object goes here
End Sub

Public Sub Delete()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
End Sub

Public Property Get IsDeleted() As Boolean
  IsDeleted = mudtProps.IsDeleted
End Property

Public Property Get IsNew() As Boolean
  IsNew = mudtProps.IsNew
End Property

Public Property Get IsDirty() As Boolean
  IsDirty = mudtProps.IsDirty
End Property

Private Function GetState() As String
  Dim udtData As TaskData
  
  LSet udtData = mudtProps
  GetState = udtData.Buffer
End Function

Private Sub SetState(Buffer As String)
  Dim udtData As TaskData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData
End Sub

Public Property Get ID() As Long
  ID = mudtProps.ID
End Property

Public Property Get Name() As String
  Name = Trim$(mudtProps.Name)
End Property

Public Property Let Name(Value As String)
  If mcolStack.Count = 0 Then Err.Raise 383
  
  mudtProps.Name = Value
  mudtProps.IsDirty = True
End Property

Public Property Get ProjectedDays() As Long
  ProjectedDays = mudtProps.ProjectedDays
End Property

Public Property Let ProjectedDays(Value As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
  If Not mudtProps.IsNew Then Err.Raise 383
  
  mudtProps.ProjectedDays = Value
  mudtProps.IsDirty = True
End Property

Public Property Get PercentComplete() As Single
  PercentComplete = mudtProps.PercentComplete
End Property

Public Property Let PercentComplete(Value As Single)
  If mcolStack.Count = 0 Then Err.Raise 383
  
  mudtProps.PercentComplete = Value
  mudtProps.IsDirty = True
End Property

Friend Function GetSuperState() As String
  GetSuperState = GetState
End Function

Friend Sub SetSuperState(Buffer As String)
  SetState Buffer
End Sub
