VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomerDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type DisplayProps
  CustomerID As Long
  Name As String * 50
  Phone As String * 25
End Type

Private mudtProps As DisplayProps

Public Property Get CustomerID() As Long

  CustomerID = mudtProps.CustomerID

End Property

Public Property Get Name() As String

  Name = Trim$(mudtProps.Name)

End Property

Public Property Get Phone() As String

  Phone = Trim$(mudtProps.Phone)

End Property

Friend Property Let CustomerID(Value As Long)

  mudtProps.CustomerID = Value

End Property

Friend Property Let Name(Value As String)

  mudtProps.Name = Value

End Property

Friend Property Let Phone(Value As String)

  mudtProps.Phone = Value

End Property

