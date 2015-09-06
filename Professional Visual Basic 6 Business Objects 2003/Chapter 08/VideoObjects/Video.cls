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

Private Type VideoProps
  VideoID As Long
  Title As String * 30
  ReleaseDate As Variant
  Studio As String * 30
  Category As String * 20
  Rating As String * 5
End Type

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

  If Not mflgEditing Then Err.Raise 445

  If mflgDeleted And Not mflgNew Then
    ' code to delete the object's data goes here
    DeleteObject mudtProps.VideoID
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

  mobjTapes.ApplyEdit mudtProps.VideoID
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

  If mflgEditing Then Err.Raise 445
  If Not mflgNew Then Err.Raise 445
  
 ' code to load the object goes here
  Fetch VideoID
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

  ReleaseDate = mudtProps.ReleaseDate

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

Private Sub Save()

  Dim rsVideo As Recordset
  Dim strSQL As String

  Set rsVideo = New Recordset
  strSQL = "SELECT * FROM Video WHERE VideoID=" & VideoID
  rsVideo.Open strSQL, cnVideo, , adLockOptimistic
  If mflgNew Then rsVideo.AddNew

  With rsVideo
    .Fields("Title") = mudtProps.Title
    .Fields("ReleaseDate") = mudtProps.ReleaseDate
    .Fields("Studio") = mudtProps.Studio
    .Fields("Rating") = mudtProps.Rating
    .Fields("Category") = mudtProps.Category
    .Update
    If mflgNew Then mudtProps.VideoID = .Fields("VideoID")
    .Close
  End With

  Set rsVideo = Nothing

End Sub

Private Sub DeleteObject(VideoID As Long)

  cnVideo.Execute "DELETE FROM Video WHERE VideoID=" & VideoID

End Sub

Private Sub Fetch(VideoID As Long)

  Dim rsVideo As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT * FROM Video WHERE VideoID=" & VideoID
  Set rsVideo = New Recordset
  rsVideo.Open strSQL, cnVideo

  With rsVideo
    mudtProps.VideoID = .Fields("VideoID")
    mudtProps.Title = .Fields("Title")
    mudtProps.ReleaseDate = .Fields("ReleaseDate")
    mudtProps.Studio = .Fields("Studio")
    mudtProps.Rating = .Fields("Rating")
    mudtProps.Category = .Fields("Category")

    mobjValid.RuleBroken "Title", False
    
    rsVideo.Close
  End With

  Set rsVideo = Nothing

End Sub

