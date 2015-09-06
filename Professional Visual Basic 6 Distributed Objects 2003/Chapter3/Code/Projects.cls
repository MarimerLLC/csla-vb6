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
Private mcolDeleted As Collection

Private mlngEditing As Long

Private Sub Class_Initialize()
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
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
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set objProject = New Project
  With objProject
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
  Set Add = objProject
  Set objProject = Nothing
End Function

Friend Sub AddProject(Child As Project)
  mcolItems.Add Child
End Sub

Public Sub Remove(ByVal Index As Variant)
  If mlngEditing = 0 Then Err.Raise 445

  With mcolItems(Index)
    .BeginEdit
    .Delete
    .ApplyEdit
  End With
  mcolDeleted.Add mcolItems(Index)
  mcolItems.Remove Index
End Sub

Public Function IsDirty() As Boolean
  Dim objProject As Project
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each objProject In mcolItems
      If objProject.IsDirty Or objProject.IsNew Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
  Dim objProject As Project

  For Each objProject In mcolItems
    objProject.BeginEdit
  Next
  For Each objProject In mcolDeleted
    objProject.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub CancelEdit()
  Dim lngindex As Long
  Dim objProject As Project

  If mlngEditing = 0 Then Err.Raise 445
  
  For lngindex = 1 To mcolItems.Count
    Set objProject = mcolItems(lngindex)
    With objProject
      If Not .IsNew Then
        .CancelEdit
      Else
        mcolItems.Remove lngindex
      End If
    End With
    Set objProject = Nothing
  Next
  For lngindex = mcolDeleted.Count To 1 Step -1
    Set objProject = mcolDeleted(lngindex)
    With objProject
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add objProject
          mcolDeleted.Remove lngindex
        End If
      Else
        mcolDeleted.Remove lngindex
      End If
    End With
    Set objProject = Nothing
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub ApplyEdit()
  Dim objProject As Project

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each objProject In mcolItems
    objProject.ApplyEdit
  Next
  For Each objProject In mcolDeleted
    objProject.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub Delete()
  Dim lngindex As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  For lngindex = mcolItems.Count To 1 Step -1
    Remove lngindex
  Next
End Sub

Friend Sub SetState(Buffer As String)
  Dim objBuffer As Buffer
  Dim lngindex As Long
  Dim lngEdit As Long
  Dim objProject As Project
  
  Set objBuffer = New Buffer
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  With objBuffer
    .SetState Buffer
    For lngindex = 1 To .Count
      Set objProject = New Project
      objProject.SetState .Item(lngindex)
      For lngEdit = 1 To mlngEditing
        objProject.BeginEdit
      Next
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





