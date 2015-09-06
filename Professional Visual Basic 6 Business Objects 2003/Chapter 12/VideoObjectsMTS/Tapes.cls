VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolTapes As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

  Set mcolTapes = New Collection

End Sub

Public Function Item(ByVal Index As Variant) As Tape
Attribute Item.VB_UserMemId = 0

  Set Item = mcolTapes(Index)

End Function

Public Function Count() As Long

  Count = mcolTapes.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

  Set NewEnum = mcolTapes.[_NewEnum]

End Function

Public Sub BeginEdit()

  Dim objTape As Tape
  
  If mflgEditing Then Err.Raise 445
  
  For Each objTape In mcolTapes
    objTape.ChildBeginEdit
  Next

  mflgEditing = True

End Sub

Public Sub CancelEdit()

  Dim objTape As Tape
  Dim lngIndex As Long
  
  If Not mflgEditing Then Err.Raise 445
  
  mflgEditing = False
  For lngIndex = mcolTapes.Count To 1 Step -1
    Set objTape = mcolTapes.Item(lngIndex)
    With objTape
      If .IsNew Then _
         mcolTapes.Remove lngIndex
      .ChildCancelEdit
    End With

  Next lngIndex

End Sub

Public Sub ApplyEdit(VideoID As Long)

  'Dim objPersist As TapesPersist
  Dim objTape As Tape
  
  If Not mflgEditing Then Err.Raise 445
  
  'If IsDirty Then
  '  Set objPersist = CreateObject("VideoServerMTS.TapesPersist", PERSIST_SERVER)
  '  SetState objPersist.Save(VideoID, GetState)
  '  Set objPersist = Nothing
  
  'Else
    For Each objTape In mcolTapes
      objTape.ChildApplyEdit (VideoID)
    Next

  'End If
  
  mflgEditing = False

End Sub

Public Sub Delete()

  Dim objTape As Tape
  
  If Not mflgEditing Then Err.Raise 445
  ' code to delete the object's data goes here

  For Each objTape In mcolTapes
    With objTape
      .BeginEdit
      .Delete
      .ApplyEdit
    End With

  Next

End Sub

Public Function Add() As Tape

  Dim objTape As Tape
  
  If Not mflgEditing Then Err.Raise 445

  Set objTape = New Tape
  objTape.SetAsChild Me
  objTape.ChildBeginEdit
  Set Add = objTape
  Set objTape = Nothing

End Function

Friend Sub AddTape(Child As Tape)

  mcolTapes.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

  If Not mflgEditing Then Err.Raise 445

  With mcolTapes(Index)
    .BeginEdit
    .Delete
    .ApplyEdit
  End With

End Sub

Friend Sub Load(VideoID As Long)

  Dim objPersist As TapesPersist
  
  Set objPersist = CreateObject("VideoServerMTS.TapesPersist", PERSIST_SERVER)
  SetState objPersist.Fetch(VideoID)
  Set objPersist = Nothing

End Sub

Friend Function GetState() As String

  Dim objBuffer As Buffer
  Dim objTape As Tape
  
  Set objBuffer = New Buffer
  Set objTape = New Tape
  objBuffer.Initialize Len(objTape.GetState), 20

  For Each objTape In mcolTapes
    objBuffer.Add objTape.GetState
  Next

  GetState = objBuffer.GetState
  Set objBuffer = Nothing

End Function

Friend Sub SetState(Buffer As String)

  Dim lngIndex As Long
  Dim objTape As Tape
  Dim objBuffer As Buffer

  Set mcolTapes = Nothing
  Set mcolTapes = New Collection

  Set objBuffer = New Buffer
  objBuffer.SetState Buffer

  For lngIndex = 1 To objBuffer.Count
    Set objTape = New Tape
    With objTape
      .SetState objBuffer.Item(lngIndex)
      .SetAsChild Me
    End With
    mcolTapes.Add Item:=objTape
  Next

End Sub

Friend Property Get IsDirty() As Boolean

  Dim objTape As Tape
  
  IsDirty = False

  For Each objTape In mcolTapes
    If objTape.IsDirty Or objTape.IsNew Then
      IsDirty = True
      Exit For
    End If

  Next

End Property

