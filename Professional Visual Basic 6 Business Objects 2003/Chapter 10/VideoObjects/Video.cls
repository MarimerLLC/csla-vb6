VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Video"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private mudtProps As VideoProps
Private mudtSave As VideoProps

Private mobjCategories As TextList
Private mobjRatings As TextList
Private mobjTapes As Tapes

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
  mobjTapes.BeginEdit

  mflgEditing = True

End Sub

Public Sub CancelEdit()

  If Not mflgEditing Then Err.Raise 445
  
  mflgEditing = False
  mflgDeleted = False
  mflgDirty = False
  ' restore object state
  LSet mudtProps = mudtSave
  mobjTapes.CancelEdit

End Sub

Public Sub ApplyEdit()

  Dim objPersist As VideoPersist
  
  If Not mflgEditing Then Err.Raise 445

  Set objPersist = CreateObject("VideoServer.VideoPersist", PERSIST_SERVER)
  
  If mflgDeleted And Not mflgNew Then
    ' code to delete the object's data goes here
    objPersist.DeleteObject mudtProps.VideoID
    mflgNew = True
    mflgDeleted = False

  ElseIf mflgDirty Or mflgNew Then
    If Not IsValid Then Err.Raise 445
    ' save object to database if appropriate
    SetState objPersist.Save(GetState)
    ' save object state
    LSet mudtSave = mudtProps
    mflgNew = False

  End If

  mobjTapes.ApplyEdit mudtProps.VideoID
  Set objPersist = Nothing
  mflgDirty = False
  mflgEditing = False

End Sub

Private Sub Class_Initialize()

  mflgNew = True
  Set mobjValid = New BrokenRules
 
  Set mobjCategories = New TextList
  mobjCategories.Load "Categories"
  mudtProps.Category = mobjCategories.Key(mobjCategories(1))
  
  Set mobjRatings = New TextList
  mobjRatings.Load "Ratings"
  mudtProps.Rating = mobjRatings.Key(mobjRatings(1))
  
  Set mobjTapes = New Tapes

 ' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  '
  mobjValid.RuleBroken "Title", True

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

Public Sub Load(VideoID As Long)

  Dim objPersist As VideoPersist
  
  If mflgEditing Then Err.Raise 445
  If Not mflgNew Then Err.Raise 445
  
 ' code to load the object goes here
  Set objPersist = CreateObject("VideoServer.VideoPersist", PERSIST_SERVER)
  SetState objPersist.Fetch(VideoID)
  Set objPersist = Nothing
    
  mobjValid.RuleBroken "Title", False
   
  mobjTapes.Load mudtProps.VideoID
 
  mflgNew = False

End Sub

Public Sub Delete()

  If Not mflgEditing Then Err.Raise 445
  
  mflgDeleted = True
  mflgDirty = True
  
  mobjTapes.Delete
  
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

Public Property Get VideoID() As Long

  VideoID = mudtProps.VideoID

End Property

Public Property Let Title(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtProps.Title) Then _
    Err.Raise vbObjectError + 1001, "String value too long"
  mobjValid.RuleBroken "Title", (Len(Trim$(Value)) = 0)
  
  mudtProps.Title = Value
  mflgDirty = True

End Property

Public Property Get Title() As String

  Title = Trim$(mudtProps.Title)

End Property

Public Property Let Studio(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtProps.Studio) Then _
    Err.Raise vbObjectError + 1001, "String value too long"
  
  mudtProps.Studio = Value
  mflgDirty = True

End Property

Public Property Get Studio() As String

  Studio = Trim$(mudtProps.Studio)

End Property

Public Property Let ReleaseDate(Value As Variant)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Trim(Value)) = 0 Then
    mudtProps.ReleaseDate = ""
    mobjValid.RuleBroken "ReleaseDate", False

  ElseIf IsDate(Value) Then
    mudtProps.ReleaseDate = CVDate(Value)
    mobjValid.RuleBroken "ReleaseDate", False

  Else
    mobjValid.RuleBroken "ReleaseDate", True

 End If

  mflgDirty = True

End Property

Public Property Get ReleaseDate() As Variant

  With mudtProps
    ReleaseDate = IIf(.ReleaseDate = 0, "", .ReleaseDate)
  End With

End Property

Public Property Let Category(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  mudtProps.Category = mobjCategories.Key(Value)
  mflgDirty = True

End Property

Public Property Get Category() As String

  Category = mobjCategories.Item(Trim$(mudtProps.Category))

End Property

Public Property Get Categories() As TextList

  Set Categories = mobjCategories

End Property

Public Property Let Rating(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  mudtProps.Rating = mobjRatings.Key(Value)
  mflgDirty = True

End Property

Public Property Get Rating() As String

  Rating = mobjRatings.Item(Trim$(mudtProps.Rating))

End Property

Public Property Get Ratings() As TextList

  Set Ratings = mobjRatings

End Property

Public Property Get Tapes() As Tapes

  Set Tapes = mobjTapes

End Property

Private Function GetState() As String

  Dim udtData As VideoData
    
  With mudtProps
    .IsDeleted = mflgDeleted
    .IsNew = mflgNew
    .IsDirty = mflgDirty
  End With

  LSet udtData = mudtProps
  GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

  Dim udtData As VideoData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData

End Sub

