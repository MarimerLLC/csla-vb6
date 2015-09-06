VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Count() As Long
  Count = gcolItems.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = gcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As Variant
Attribute Item.VB_UserMemId = 0
  If VarType(gcolItems(Index)) = vbObject Then
    Set Item = gcolItems(Index)
  Else
    Item = gcolItems(Index)
  End If
End Function

Public Sub Add(Value As Variant, Optional Key As String)
  If Len(Key) > 0 Then
    gcolItems.Add Value, Key
  Else
    gcolItems.Add Value
  End If
End Sub

Public Function CreateObject(ProgID As String, Optional Key As String) As Object
  Dim objNew As Object
  
  Set objNew = VBA.CreateObject(ProgID)
  If Len(Key) > 0 Then
    gcolItems.Add objNew, Key
  Else
    gcolItems.Add objNew
  End If
  Set CreateObject = objNew
  Set objNew = Nothing
End Function

Public Sub Remove(ByVal Index As Variant)
  On Error Resume Next
  gcolItems.Remove Index
End Sub

