VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Projects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolItems As Collection

Private mflgEditing As Boolean

Private Sub Class_Initialize()
  Set mcolItems = New Collection
End Sub

Public Function Count() As Long
  Count = mcolItems.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As Project
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function

Public Function Add() As Project
  Dim objProject As Project
  
  If Not mflgEditing Then Err.Raise 445
  
  Set objProject = New Project
  objProject.Initialize Me
  objProject.ChildBeginEdit
  Set Add = objProject
  Set objProject = Nothing
End Function

Friend Sub AddProject(Child As Project)
  mcolItems.Add Child
End Sub

Public Sub Remove(ByVal Index As Variant)
  If Not mflgEditing Then Err.Raise 445

  With mcolItems(Index)
    .BeginEdit
    .Delete
    .ApplyEdit
  End With
End Sub

Public Sub Delete()
  Dim objProject As Project
  
  If Not mflgEditing Then Err.Raise 445
  
  For Each objProject In mcolItems
    With objProject
      .BeginEdit
      .Delete
      .ApplyEdit
    End With
  Next
End Sub

Friend Sub SetState(Buffer As String)
  Dim objBuffer As Buffer
  Dim lngIndex As Long
  Dim objProject As Project
  
  Set objBuffer = New Buffer
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  With objBuffer
    .SetState Buffer
    For lngIndex = 1 To .Count
      Set objProject = New Project
      objProject.SetState .Item(lngIndex)
      If mflgEditing Then objProject.ChildBeginEdit
      mcolItems.Add objProject
      Set objProject = Nothing
    Next
  End With
End Sub

Friend Function GetState() As String
  Dim objBuffer As Buffer
  Dim objProject As Project
  Dim udtData As ProjectData
  
  Set objBuffer = New Buffer
  With objBuffer
    .Initialize Len(udtData.Buffer), 10
    For Each objProject In mcolItems
      .Add objProject.GetState
    Next
    GetState = .GetState
  End With
  Set objBuffer = Nothing
End Function

Public Sub BeginEdit()
  Dim objProject As Project

  If mflgEditing Then Err.Raise 445
  
  For Each objProject In mcolItems
    objProject.ChildBeginEdit
  Next
  
  mflgEditing = True
End Sub

Public Sub ApplyEdit()
  Dim objProject As Project

  If Not mflgEditing Then Err.Raise 445
  
  For Each objProject In mcolItems
    objProject.ChildApplyEdit
  Next
  
  mflgEditing = False
End Sub

Public Sub CancelEdit()
  Dim lngIndex As Long
  Dim objProject As Project

  If Not mflgEditing Then Err.Raise 445
  
  For lngIndex = 1 To mcolItems.Count
    Set objProject = mcolItems(lngIndex)
    With objProject
      If Not .IsNew Then
        .ChildCancelEdit
      Else
        mcolItems.Remove lngIndex
      End If
    End With
    Set objProject = Nothing
  Next
  
  mflgEditing = False
End Sub

Public Function IsDirty() As Boolean
  Dim objProject As Project
  
  For Each objProject In mcolItems
    If objProject.IsDirty Or objProject.IsNew Then
      IsDirty = True
      Exit For
    End If
  Next
End Function

