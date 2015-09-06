VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private mudtProps As ClientProps
Private mudtSave As ClientData
Private mcolProjects As Projects

Private mflgEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit()
  If mflgEditing Then Err.Raise 445
  
  ' save object state
  LSet mudtSave = mudtProps
  mflgEditing = True
  mcolProjects.BeginEdit
End Sub

Public Sub CancelEdit()
  If Not mflgEditing Then Err.Raise 445
  
  mflgEditing = False
  mudtProps.IsDeleted = False
  ' restore object state
  LSet mudtProps = mudtSave
  mcolProjects.CancelEdit
End Sub

Public Sub ApplyEdit()
  Dim objPersist As ClientPersist
  Dim strProjects As String
  
  If Not mflgEditing Then Err.Raise 445

  ' Set objPersist = CreateObject("TaskServer.ClientPersist")
  Set objPersist = New ClientPersist
  If mudtProps.IsDeleted Then
    ' code to delete the object's data goes here
    objPersist.DeleteObject mudtProps.ID
    mudtProps.IsNew = True
    mudtProps.IsDeleted = False
  ElseIf IsDirty Or mudtProps.IsNew Then
    If Not IsValid Then Err.Raise 445
    ' save object to database if appropriate
    ' save object state
    LSet mudtSave = mudtProps
    strProjects = mcolProjects.GetState
    SetState objPersist.Save(GetState, strProjects)
    mcolProjects.SetState strProjects
    mudtProps.IsNew = False
  End If
  mudtProps.IsDirty = False
  mflgEditing = False
  mcolProjects.ApplyEdit
End Sub

Private Sub Class_Initialize()
  Set mobjValid = New BrokenRules
  Set mcolProjects = New Projects
  
  mudtProps.IsNew = True
  ' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  '
  ' mobjValid.RuleBroken "RuleName", True
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

Public Sub Load(ID As Long)
  Dim objPersist As ClientPersist
  Dim strProjects As String
  Dim strTasks As String
  
  If mflgEditing Then Err.Raise 445
  If Not mudtProps.IsNew Then Err.Raise 445
  
  mudtProps.IsNew = False

  ' code to load the object goes here
  Set objPersist = New ClientPersist
  SetState objPersist.Fetch(ID, strProjects)
  Set objPersist = Nothing
  Set mcolProjects = New Projects
  mcolProjects.SetState strProjects
End Sub

Public Sub Delete()
  If Not mflgEditing Then Err.Raise 445
  
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
  IsDirty = mudtProps.IsDirty Or mcolProjects.IsDirty
End Property

Private Function GetState() As String
  Dim udtData As ClientData
  
  LSet udtData = mudtProps
  GetState = udtData.Buffer
End Function

Private Sub SetState(Buffer As String)
  Dim udtData As ClientData
  
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
  If Not mflgEditing Then Err.Raise 383
  
  mudtProps.Name = Value
  mudtProps.IsDirty = True
End Property

Public Property Get ContactName() As String
  ContactName = Trim$(mudtProps.ContactName)
End Property

Public Property Let ContactName(Value As String)
  If Not mflgEditing Then Err.Raise 383
  
  mudtProps.ContactName = Value
  mudtProps.IsDirty = True
End Property

Public Property Get Phone() As String
  Phone = Trim$(mudtProps.Phone)
End Property

Public Property Let Phone(Value As String)
  If Not mflgEditing Then Err.Raise 383
  
  mudtProps.Phone = Value
  mudtProps.IsDirty = True
End Property

Public Property Get Projects() As Projects
  Set Projects = mcolProjects
End Property

