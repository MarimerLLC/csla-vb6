VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Public Function Item(ByVal Index As Variant) As Task
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function

Public Function Add() As Task
  Dim objTask As Task
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set objTask = New Task
  With objTask
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
  Set Add = objTask
  Set objTask = Nothing
End Function

Friend Sub AddTask(Child As Task)

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
  Dim objTask As Task
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each objTask In mcolItems
      If objTask.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
  Dim objTask As Task

  For Each objTask In mcolItems
    objTask.BeginEdit
  Next
  For Each objTask In mcolDeleted
    objTask.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub ApplyEdit()
  Dim objTask As Task

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each objTask In mcolItems
    objTask.ApplyEdit
  Next
  For Each objTask In mcolDeleted
    objTask.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub CancelEdit()
  Dim lngindex As Long
  Dim objTask As Task

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each objTask In mcolItems
    objTask.CancelEdit
  Next
  For lngindex = mcolDeleted.Count To 1 Step -1
    Set objTask = mcolDeleted(lngindex)
    With objTask
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add objTask
          mcolDeleted.Remove lngindex
        End If
      Else
        mcolDeleted.Remove lngindex
      End If
    End With
    Set objTask = Nothing
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

Friend Function GetSuperState() As String
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim objTask As Task
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each objTask In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), objTask.GetSuperState
    Next
    For Each objTask In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), objTask.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
End Function

Friend Sub SetSuperState(Buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngindex As Long
  Dim lngEdit As Long
  Dim objTask As Task
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = Buffer
    .Contents = arBuffer
    For lngindex = 1 To .ReadProperty("Count")
      Set objTask = New Task
      objTask.SetSuperState .ReadProperty("Item" & CStr(lngindex))
      For lngEdit = 1 To mlngEditing
        objTask.BeginEdit
      Next
      If Not objTask.IsDeleted Then
        mcolItems.Add objTask
      Else
        mcolDeleted.Add objTask
      End If
      Set objTask = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub







