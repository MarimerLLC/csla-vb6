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

  Dim objTape As Tape
  Dim lngIndex As Long

  If Not mflgEditing Then Err.Raise 445

  For lngIndex = mcolTapes.Count To 1 Step -1
    Set objTape = mcolTapes.Item(lngIndex)
    With objTape
      If .IsDeleted Then _
          mcolTapes.Remove lngIndex
      .ChildApplyEdit VideoID
    End With

  Next lngIndex

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

  Dim rsTape As Recordset
  Dim strSQL As String
  Dim objTape As Tape

  If mflgEditing Then Err.Raise 445
  
   strSQL = "SELECT TapeID FROM Tape WHERE VideoID=" & VideoID
  Set rsTape = New Recordset
  rsTape.Open strSQL, cnVideo

  Do While Not rsTape.EOF
    Set objTape = New Tape

    With objTape
      .ChildLoad rsTape("TapeID")
      .SetAsChild Me
    End With

    mcolTapes.Add Item:=objTape
    rsTape.MoveNext

  Loop

  rsTape.Close
  Set rsTape = Nothing

End Sub

