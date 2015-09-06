VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusinessObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Count() As Long
  Count = gcolObjects.Count
End Function

Public Sub Add(Object As Object, Key As String)
  gcolObjects.Add Object, Key
End Sub

Public Function Item(Key As String) As Object
  Set Item = gcolObjects.Item(Key)
End Function

Public Sub Remove(Key As String)
  gcolObjects.Remove Key
End Sub

Public Sub RemoveAll()
  With gcolObjects
    Do While .Count > 0
      .Remove 1
    Loop
  End With
End Sub

