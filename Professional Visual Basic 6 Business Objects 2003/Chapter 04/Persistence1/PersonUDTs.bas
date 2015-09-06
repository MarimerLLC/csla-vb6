Attribute VB_Name = "PersonUDTs"
Option Explicit

Public Type PersonProps
  SSN As String * 11
  Name As String * 50
  Birthdate As Date
End Type

Public Type PersonData

  Buffer As String * 68

End Type

