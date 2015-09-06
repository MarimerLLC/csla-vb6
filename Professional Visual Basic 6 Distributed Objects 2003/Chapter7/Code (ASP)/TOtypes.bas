Attribute VB_Name = "TOtypes"
Option Explicit

Public Type ClientProps
  IsNew As Boolean
  IsDirty As Boolean
  IsDeleted As Boolean
  ID As Long
  Name As String * 50
  ContactName As String * 50
  Phone As String * 25
End Type

Public Type ClientData
  Buffer As String * 131
End Type

Public Type ProjectProps
  IsNew As Boolean
  IsDirty As Boolean
  IsDeleted As Boolean
  ID As Long
  Name As String * 50
End Type

Public Type TaskData
  Buffer As String * 60
End Type

Public Type TaskProps
  IsNew As Boolean
  IsDirty As Boolean
  IsDeleted As Boolean
  ID As Long
  Name As String * 50
  ProjectedDays As Long
  PercentComplete As Single
End Type


Public Type ProjectData
  Buffer As String * 56
End Type

