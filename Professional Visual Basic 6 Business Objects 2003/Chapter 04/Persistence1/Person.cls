VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event NewAge()
Event Valid(IsValid As Boolean)

Private mintAge As Integer
Private mudtPerson As PersonProps
Private mudtSaved As PersonProps
Private mflgEditing As Boolean
Private mflgNew As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
 
Public Property Let SSN(Value As String)

  If Not mflgEditing Then Err.Raise 383
  If Not mflgNew Then Err.Raise 383

  If Len(Value) > 11 Then _
    Err.Raise vbObjectError + 1001, "Person", "SSN too long"

  mudtPerson.SSN = Value
  mobjValid.RuleBroken "SSN", (Len(Trim$(mudtPerson.SSN)) <> 11)
  
End Property

Public Property Get SSN() As String

  SSN = Trim$(mudtPerson.SSN)

End Property

Public Property Let Name(Value As String)

  If Not mflgEditing Then Err.Raise 383
  mudtPerson.Name = Value

End Property

Public Property Get Name() As String

  Name = Trim$(mudtPerson.Name)

End Property

Public Property Let Birthdate(Value As Date)

  Static intOldAge As Integer
  
  If Not mflgEditing Then Err.Raise 445
  mudtPerson.Birthdate = Value

  CalculateAge

  If mintAge <> intOldAge Then
    intOldAge = mintAge
    RaiseEvent NewAge

  End If

End Property

Public Property Get Birthdate() As Date

  Birthdate = mudtPerson.Birthdate

End Property

Public Property Get Age() As Integer

  Age = mintAge

End Property

Private Sub CalculateAge()

  If DatePart("y", mudtPerson.Birthdate) > DatePart("y", Now) Then
    mintAge = DateDiff("yyyy", mudtPerson.Birthdate, Now) - 1

  Else
 
    mintAge = DateDiff("yyyy", mudtPerson.Birthdate, Now)

  End If

End Sub

Private Sub Class_Initialize()

  mflgEditing = False
  mflgNew = True
  
  Set mobjValid = New BrokenRules
  mobjValid.RuleBroken "SSN", True

End Sub

Private Sub mobjValid_BrokenRule()

  RaiseEvent Valid(False)

End Sub

Private Sub mobjValid_NoBrokenRules()

  RaiseEvent Valid(True)

End Sub

Public Property Get IsValid() As Boolean

  IsValid = (mobjValid.Count = 0)

End Property

Public Sub BeginEdit()

  LSet mudtSaved = mudtPerson
  mflgEditing = True

End Sub

Public Sub ApplyEdit()

' data would be saved here
  mflgEditing = False
  mflgNew = False
  
End Sub

Public Sub CancelEdit()

  If Not mflgEditing Then Err.Raise 445
  LSet mudtPerson = mudtSaved
  mflgEditing = False

End Sub

Public Property Get IsNew() As Boolean
 
  IsNew = mflgNew

End Property

Public Function GetState() As String

  Dim udtBuffer As PersonData
  
  LSet udtBuffer = mudtPerson
  GetState = udtBuffer.Buffer

End Function

Public Sub SetState(ByVal Buffer As String)

  Dim udtBuffer As PersonData
  
  udtBuffer.Buffer = Buffer
  LSet mudtPerson = udtBuffer
  CalculateAge

End Sub

