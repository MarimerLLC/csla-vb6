VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BrokenRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event BrokenRule()
Event NoBrokenRules()

Private colBroken As Collection


Private Sub Class_Initialize()
  Set colBroken = New Collection
End Sub

Public Sub RuleBroken(Rule As String, IsBroken As Boolean)
  On Error GoTo HandleError
  If IsBroken Then
    colBroken.Add True, Rule
    RaiseEvent BrokenRule
  Else
    colBroken.Remove Rule
    If colBroken.Count = 0 Then RaiseEvent NoBrokenRules
  End If

HandleError:
End Sub

Public Property Get Count() As Integer
  Count = colBroken.Count
End Property


