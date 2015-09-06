Attribute VB_Name = "VideoTypes"
Option Explicit

Public Type CustomerProps
  CustomerID As Long
  Name As String * 50
  Phone As String * 25
  Address1 As String * 30
  Address2 As String * 30
  City As String * 20
  State As String * 2
  ZipCode As String * 10
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type CustomerData
  Buffer As String * 172
End Type

Public Type TextListProps
  Key As String * 30
  Item As String * 255
End Type

Public Type TextListData
  Buffer As String * 285
End Type

Public Type CustDisplayProps
  CustomerID As Long
  Name As String * 50
  Phone As String * 25
End Type

Public Type CustDisplayData
  Buffer As String * 78
End Type

Public Type VideoDisplayProps
  VideoID As Long
  Title As String * 30
  ReleaseDate As Date
End Type

Public Type VideoDisplayData
  Buffer As String * 36
End Type

Public Type VideoProps
  VideoID As Long
  Title As String * 30
  ReleaseDate As Date
  Studio As String * 30
  Category As String * 20
  Rating As String * 5
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type VideoData
  Buffer As String * 94
End Type

Public Type TapeProps
  TapeID As Long
  VideoID As Long
  Title As String * 30
  CheckedOut As Boolean
  DateAcquired As Date
  DateDue As Date
  LateFee As Boolean
  InvoiceID As Long
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type TapeData
  Buffer As String * 52
End Type

Public Type InvoiceProps
  InvoiceID As Long
  CustomerID As Long
  CustomerName As String * 50
  CustomerPhone As String * 25
  SubTotal As Double
  Tax As Double
  Total As Double
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type InvoiceData
  Buffer As String * 96
End Type

Public Type FeeProps
  InvoiceID As Long
  FeeID As Long
  VideoTitle As String * 30
  EnteredDate As Date
  DaysOver As Integer
  Paid As Boolean
  PaidDate As Date
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type FeeData
  Buffer As String * 48
End Type

Public Type InvoiceTapeProps
  InvoiceID As Long
  ItemID As Long
  TapeID As Long
  Title As String * 30
  Price As Double
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type InvoiceTapeData
  Buffer As String * 44
End Type


